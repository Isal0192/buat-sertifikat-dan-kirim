import os
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
from pathlib import Path
import logging
import re
from dotenv import load_dotenv

class CertificateGenerator:
    def __init__(self, excel_path: str, template_path: str, output_dir: str):
        
        self.excel_path = Path(excel_path)
        self.template_path = Path(template_path)
        self.output_dir = Path(output_dir)
        self.logger = self._setup_logging()

    def _setup_logging(self) -> logging.Logger:
        
        logger = logging.getLogger(__name__)
        logger.setLevel(logging.INFO)

        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)

        return logger

    def validate_paths(self) -> None:
        """Validate Excel and template file paths."""
        if not self.excel_path.exists() or self.excel_path.suffix != '.xlsx':
            raise ValueError("Excel file tidak di temkan")
        if not self.template_path.exists() or self.template_path.suffix != '.docx':
            raise ValueError("templat file tidak di temukan")

    def read_excel(self) -> pd.DataFrame:
        """baca file excel."""
        try:
            df = pd.read_excel(self.excel_path)
            if df.empty:
                raise ValueError("file excel kosong")
            
            # Log available columns in the Excel file for debugging
            self.logger.info(f"column tersedia dalam file excel : {df.columns.tolist()}. 👍🏻")
            return df
        except Exception as e:
            self.logger.error(f"gagal membaca file excel: {e}")
            raise

    @staticmethod
    def validate_email(email: str) -> bool:
        """Vvalidasi email ."""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email))

    @staticmethod
    def format_name(name: str) -> str:
        """format nama."""
        if not name or not isinstance(name, str):
            raise ValueError("format nama salah")
        
        formatted = ' '.join(name.strip().split()[:3])
        return ' '.join(word.title() for word in formatted.split())

    def create_certificate(self, name: str) -> Path:
        """Create individual certificate PDF."""
        try:
            doc = DocxTemplate(self.template_path)
            doc.render({"NAMA": name})
            
            # Create output directory if it doesn't exist
            self.output_dir.mkdir(parents=True, exist_ok=True)
            
            docx_path = self.output_dir / f"{name}.docx"
            pdf_path = self.output_dir / f"{name}.pdf"
            
            doc.save(docx_path)
            convert(docx_path)
            docx_path.unlink()  # Remove temporary .docx file
            
            return pdf_path
        except Exception as e:
            self.logger.error(f"gagal membuat sertifikat untuk {name}: {e}")
            raise

    def generate_certificates(self, name_column: str, email_column: str = None):
        """Generate certificates for all participants."""
        df = self.read_excel()

        # Cek apakah kolom nama ada dalam DataFrame
        if name_column not in df.columns:
            raise ValueError(f"column tidak ada: {name_column}")
        
        # Jika email_column tidak None, periksa apakah kolom email ada
        if email_column and email_column not in df.columns:
            raise ValueError(f"colum tidak ada: {email_column}")

        # Validasi dan proses data
        pdf_paths = []
        valid_emails = []
        formatted_names = []
        
        for _, row in df.iterrows():
            name = row[name_column]
            email = row[email_column] if email_column else None  # Jika email_column None, abaikan email
            
            try:
                # Jika email_column ada, lakukan validasi email
                if email_column and not self.validate_email(str(email)):
                    self.logger.warning(f"format email salah: {email}")
                    continue
                
                formatted_name = self.format_name(str(name))
                pdf_path = self.create_certificate(formatted_name)
                
                pdf_paths.append(pdf_path)
                if email_column:  # Hanya masukkan email jika email_column ada
                    valid_emails.append(email)
                formatted_names.append(formatted_name)
            
            except Exception as e:
                self.logger.error(f"gagal memperoses data  {name}: {e}")
                continue
        
        return pdf_paths, valid_emails, formatted_names

def main():
    # Load environment variables
    load_dotenv()
    
    # Get email configuration from environment variables
    EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')
    EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
    SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
    SMTP_PORT = int(os.getenv('SMTP_PORT', '465'))
    
    if not all([EMAIL_ADDRESS, EMAIL_PASSWORD]):
        raise ValueError("email konfigurasi tidak di temukan dalam environment variable")

    try:
        excel_path = input("masukan path file excel seminar (e.g., pendaftaran.xlsx/absen.xlsx): ").strip()
        template_path = input("masukan path template sertifikat (e.g., template.docx): ").strip()
        
        generator = CertificateGenerator(excel_path, template_path, output_dir=f"sertifikat/sertifikat_{excel_path}")
        generator.validate_paths()
        
        df = generator.read_excel()
        print("\nColumns yang tersedia di excel:")
        print(df.columns.tolist())
        
        name_column = input("masukan column nama peserta seminar: ").strip()
        email_column = input("masukan column email peserta seminar (atau '-' jika tidak ingin mengirim email): ").strip()

        if email_column == "-":
            email_column = None
            print("Hanya membuat sertifikat, tidak mengirimkannya.")
            pdf_paths, names = generator.generate_certificates(name_column)
            print("\nSertifikat berhasil dibuat")
        else:
            email_subject = input("subject email: ").strip()
            email_body = input("pesan email: ").strip()
        
            pdf_paths, emails, names = generator.generate_certificates(name_column, email_column)
        
            from sendemail import send_email
            for pdf_path, email, name in zip(pdf_paths, emails, names):
                try:
                    send_email(
                        subject=email_subject,
                        body=email_body,
                        to=email,
                        from_email=EMAIL_ADDRESS,
                        password=EMAIL_PASSWORD,
                        smtp_server=SMTP_SERVER,
                        smtp_port=SMTP_PORT,
                        attachment=str(pdf_path)
                    )
                    print(f"sertifikat dikirim ke  {email} ({name}) oleh {os.environ['EMAIL_ADDRESS']}")
                except Exception as e:
                    print(f"Failed to send email to {email}: {e}")
            
            print("\nSertifikat berhasil dibuat dan selesai dikirim")
            
    except Exception as e:
        print(f"Errorl: {e}")

if __name__ == "__main__":
    main()
