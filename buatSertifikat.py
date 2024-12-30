# buatSertifikat.py
import os
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
from pathlib import Path
import logging
from typing import List, Tuple
import re
from dotenv import load_dotenv

class CertificateGenerator:
    def __init__(self, excel_path: str, template_path: str, output_dir: str):
        """
        Initialize the certificate generator.
        
        Args:
            excel_path: Path to Excel file containing participant data
            template_path: Path to certificate template file
            output_dir: Directory to store generated certificates
        """
        self.excel_path = Path(excel_path)
        self.template_path = Path(template_path)
        self.output_dir = Path(output_dir)
        self.logger = self._setup_logging()

    def _setup_logging(self) -> logging.Logger:
        """Configure logging for the certificate generator."""
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
            raise ValueError("Excel file not found or invalid format")
        if not self.template_path.exists() or self.template_path.suffix != '.docx':
            raise ValueError("Template file not found or invalid format")

    def read_excel(self) -> pd.DataFrame:
        """Read and validate Excel data."""
        try:
            df = pd.read_excel(self.excel_path)
            if df.empty:
                raise ValueError("Excel file is empty")
            return df
        except Exception as e:
            self.logger.error(f"Failed to read Excel file: {e}")
            raise

    @staticmethod
    def validate_email(email: str) -> bool:
        """Validate email format using regex."""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email))

    @staticmethod
    def format_name(name: str) -> str:
        """Format name with proper capitalization and length limit."""
        if not name or not isinstance(name, str):
            raise ValueError("Invalid name provided")
        
        
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

    def generate_certificates(self, name_column: str, email_column: str) -> Tuple[List[Path], List[str], List[str]]:
        """Generate certificates for all participants."""
        df = self.read_excel()
        
        if name_column not in df.columns or email_column not in df.columns:
            raise ValueError(f"Required columns not found: {name_column} or {email_column}")

        # Validate and process data
        pdf_paths = []
        valid_emails = []
        formatted_names = []
        
        for _, row in df.iterrows():
            name = row[name_column]
            email = row[email_column]
            
            try:
                if not self.validate_email(str(email)):
                    self.logger.warning(f"Invalid email format: {email}")
                    continue
                
                formatted_name = self.format_name(str(name))
                pdf_path = self.create_certificate(formatted_name)
                
                pdf_paths.append(pdf_path)
                valid_emails.append(email)
                formatted_names.append(formatted_name)
                
            except Exception as e:
                self.logger.error(f"Error processing entry {name}: {e}")
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
        raise ValueError("email configurasi tidak di temukan dalam environment variable")

    try:
        excel_path = input("masukan path file excel seminar (e.g., pendaftaran.xlsx/absen.xlsx): ").strip()
        template_path = input("masukan path template sertifikat (e.g., template.docx): ").strip()
        
        generator = CertificateGenerator(excel_path, template_path, output_dir = f"sertifikat/sertifikat_{excel_path}" )
        generator.validate_paths()
        
        df = generator.read_excel()
        print("\nColumns yang tersedia di excel:")
        print(df.columns.tolist())
        
        name_column = input("masukan column nama peserta seminar: ").strip()
        email_column = input("masukan column email peserta seminar: ").strip()
        email_subject = input("subject email: ").strip()
        email_body = input("pesan email: ").strip()
        
        pdf_paths, emails, names = generator.generate_certificates(name_column, email_column)
        
        # Import send_email only when needed
        # from sendemail import send_email
        
        # for pdf_path, email, name in zip(pdf_paths, emails, names):
        #     try:
        #         send_email(
        #             subject=email_subject,
        #             body=email_body,
        #             to=email,
        #             from_email=EMAIL_ADDRESS,
        #             password=EMAIL_PASSWORD,
        #             smtp_server=SMTP_SERVER,
        #             smtp_port=SMTP_PORT,
        #             attachment=str(pdf_path)
        #         )
        #         print(f"Certificate sent to {email} ({name})")
        #     except Exception as e:
        #         print(f"Failed to send email to {email}: {e}")
        
        # print("\nsertifikat berhasil di buat dan selesai di kirim")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()