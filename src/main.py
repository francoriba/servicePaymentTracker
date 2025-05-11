import os
import pytesseract
from PIL import Image
import PyPDF2
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import re
from pdf2image import convert_from_path

class BillProcessor:
    def __init__(self, credentials_path, receipts_folder):
        self.credentials_path = credentials_path
        self.receipts_folder = receipts_folder
        self.spreadsheet = None
        self.initialize_gspread()

    def initialize_gspread(self):
        """Inicializa la conexión con Google Sheets"""
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name(
            self.credentials_path, scope)
        client = gspread.authorize(credentials)
        self.spreadsheet = client.open('Impuestos')  # Nombre del spreadsheet

    def process_files(self):
        """Procesa todos los archivos en la carpeta de recibos"""
        for filename in os.listdir(self.receipts_folder):
            file_path = os.path.join(self.receipts_folder, filename)
            if filename.lower().endswith(('.pdf')):
                data = self.extract_from_pdf(file_path)
            elif filename.lower().endswith(('.jpg', '.jpeg', '.png')):
                data = self.extract_from_image(file_path)
            else:
                continue

            if data:
                self.update_spreadsheet(data)

    def extract_from_pdf(self, file_path):
        """Extrae información de archivos PDF"""
        try:
            with open(file_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                text = ''
                for page in reader.pages:
                    page_text = page.extract_text()
                    text += page_text

                # Si no se detecta texto, convertir PDF a imagen y usar OCR
                if not text.strip():
                    print(f"No se detectó texto en {file_path}, intentando con OCR...")
                    
                    # Convertir PDF a imágenes
                    images = convert_from_path(file_path)
                    text = ''
                    
                    # Procesar cada página con OCR
                    for image in images:
                        text += pytesseract.image_to_string(image, lang='spa') + '\n'
                
                return self.parse_bill_data(text)
        except Exception as e:
            print(f"Error procesando PDF {file_path}: {str(e)}")
            return None

    def extract_from_image(self, file_path):
        """Extrae información de imágenes usando OCR"""
        try:
            image = Image.open(file_path)
            text = pytesseract.image_to_string(image, lang='spa')
            return self.parse_bill_data(text)
        except Exception as e:
            print(f"Error procesando imagen {file_path}: {str(e)}")
            return None

    def parse_bill_data(self, text):
        """Analiza el texto extraído para obtener la información relevante"""
        # Patrones para detectar diferentes tipos de servicios
        patterns = {
            'luz': r'(luz|edenor|edesur|epec)',
            'gas': r'(gas|metrogas|naturgy|ecogas)',
            'internet': r'(internet|fibertel|telecom|movistar|personal|flow)',
            'agua': r'(agua|aysa|aguas\s*cordobesas)',
            'expensas': r'(expensas|consorcio|banco\s*roela|siro)'
        }

        # Buscar fecha (formato DD/MM/YYYY)
        date_pattern = r'\d{2}/\d{2}/\d{4}'
        date_match = re.search(date_pattern, text)
        
        # Buscar monto (formato $X.XXX,XX)
        amount_pattern = r'\$\s*[\d.,]+(?:,\d{2})?'
        amount_match = re.search(amount_pattern, text)

        # Determinar tipo de servicio
        service_type = None
        for service, pattern in patterns.items():
            if re.search(pattern, text.lower()):
                service_type = service
                break

        if date_match and amount_match and service_type:
            date = datetime.strptime(date_match.group(), '%d/%m/%Y')
            amount = amount_match.group().replace('$', '').strip()
            
            return {
                'date': date,
                'amount': amount,
                'service': service_type
            }
        return None

    def update_spreadsheet(self, data):
        """Actualiza el Google Spreadsheet con la información extraída"""
        try:
            year = str(data['date'].year)
            month = data['date'].strftime('%B')
            
            # Buscar o crear hoja para el año
            worksheet = None
            try:
                worksheet = self.spreadsheet.worksheet(year)
            except gspread.exceptions.WorksheetNotFound:
                worksheet = self.spreadsheet.add_worksheet(
                    title=year, rows=100, cols=20)
                # Agregar encabezados
                headers = ['Fecha', 'Servicio', 'Monto']
                worksheet.append_row(headers)

            # Verificar si el pago ya existe
            existing_data = worksheet.get_all_records()
            payment_date = data['date'].strftime('%d/%m/%Y')
            
            for row in existing_data:
                if (row['Fecha'] == payment_date and 
                    row['Servicio'] == data['service']):
                    print(f"Pago ya registrado: {data['service']} - {payment_date}")
                    return

            # Agregar nueva fila
            new_row = [
                payment_date,
                data['service'],
                data['amount']
            ]
            worksheet.append_row(new_row)
            print(f"Pago registrado: {data['service']} - {payment_date}")

        except Exception as e:
            print(f"Error actualizando spreadsheet: {str(e)}")

def main():
    credentials_path = '../credentials.json'
    receipts_folder = '../comprobantes'
    
    processor = BillProcessor(credentials_path, receipts_folder)
    processor.process_files()

if __name__ == "__main__":
    main()
