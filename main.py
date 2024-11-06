import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import streamlit as st
import pandas as pd
from typing import List, Optional
from dataclasses import dataclass
import logging
from pathlib import Path
import sys

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_sender.log'),
        logging.StreamHandler(sys.stdout)
    ]
)

@dataclass
class EmailConfig:
    """Classe para armazenar configurações de email"""
    smtp_user: str
    smtp_password: str
    smtp_server: str = "smtp.gmail.com"
    smtp_port: int = 587

class EmailSender:
    """Classe responsável pelo envio de emails"""
    def __init__(self, config: EmailConfig):
        self.config = config

    def send_email(self, to_email: str, subject: str, body: str) -> bool:
        """
        Envia um email individual
        Retorna True se o envio for bem sucedido, False caso contrário
        """
        try:
            msg = MIMEMultipart()
            msg['Subject'] = subject
            msg['From'] = self.config.smtp_user
            msg['To'] = to_email
            msg['Reply-To'] = 'noreply@bpalma.com'

            msg.attach(MIMEText(body, 'html'))

            with smtplib.SMTP(self.config.smtp_server, self.config.smtp_port) as server:
                server.starttls()
                server.login(self.config.smtp_user, self.config.smtp_password)
                server.sendmail(msg['From'], [msg['To']], msg.as_string())
            
            logging.info(f"Email enviado com sucesso para: {to_email}")
            return True

        except Exception as e:
            logging.error(f"Erro ao enviar email para {to_email}: {str(e)}")
            return False

class EmailProcessor:
    """Classe para processar lista de emails"""
    @staticmethod
    def validate_email(email: str) -> bool:
        """Validação básica de email"""
        return isinstance(email, str) and '@' in email and '.' in email

    @staticmethod
    def process_excel(file, email_column: str) -> List[str]:
        """Processa arquivo Excel e retorna lista de emails válidos"""
        try:
            df = pd.read_excel(file)
            emails = df[email_column].dropna().tolist()
            return [email for email in emails if EmailProcessor.validate_email(email)]
        except Exception as e:
            logging.error(f"Erro ao processar arquivo Excel: {str(e)}")
            raise

class StreamlitUI:
    """Classe para interface do usuário Streamlit"""
    def __init__(self):
        self.config = None
        
    def sidebar(self):
        """Configura a barra lateral"""
        st.sidebar.header("Configurações", divider=True)
        st.sidebar.info("Digite suas informações de email e senha do google apps.")
        smtp_user = st.sidebar.text_input("Digite seu e-mail")
        smtp_password = st.sidebar.text_input(
            "Digite sua senha", 
            placeholder="Senha apps google", 
            type="password"
        )
        
        if smtp_user and smtp_password:
            self.config = EmailConfig(smtp_user=smtp_user, smtp_password=smtp_password)
        
        self.show_help_section()

    def show_help_section(self):
        """Exibe seção de ajuda"""
        st.sidebar.header("Ajuda", divider=True)
        st.sidebar.info("Clique em Download para baixar o tutorial de como criar a senha no google apps.")
        self.download_tutorial()

        
        st.sidebar.markdown("""
       💜 Gostou do projeto? Me pague um café!
        
        **PIX:** hugorogerio522@gmail.com""")
       

    def download_tutorial(self):
        """Gerencia download do tutorial"""
        tutorial_path = Path("Como configurar uma senha.pdf")
        if tutorial_path.exists():
            with open(tutorial_path, "rb") as f:
                st.sidebar.download_button(
                    "⬇️ Download", 
                    f, 
                    file_name=tutorial_path.name
                )

    def main_page(self):
        """Página principal"""
        st.header("SISTEMA DE ENVIO DE E-MAILS EM MASSA", divider=True)
        
        subject = st.text_input("Digite o assunto do e-mail", placeholder="Assunto do e-mail")
        body = st.text_area("Digite o corpo do e-mail", height=300)
        
        single_recipient = st.text_input(
            "Digite o e-mail do destinatário único (opcional)", 
            placeholder="exemplo@dominio.com"
        )
        
        file = st.file_uploader("Ou carregue uma lista de e-mails (.xlsx)", type="xlsx")
        email_column = None

        if file:
            try:
                df = pd.read_excel(file)
                st.write("Colunas encontradas no arquivo:", list(df.columns))
                email_column = st.selectbox(
                    "Selecione a coluna que contém os emails:",
                    options=list(df.columns)
                )
            except Exception as e:
                st.error(f"Erro ao ler arquivo: {str(e)}")
                return

        if st.button("Enviar e-mail"):
            self.handle_send_button(single_recipient, file, email_column, subject, body)

    def handle_send_button(self, single_recipient, file, email_column, subject, body):
        """Processa o envio de emails"""
        if not self.config:
            st.error("Configure seu email e senha nas configurações!")
            return

        try:
            emails = self.get_email_list(single_recipient, file, email_column)
            if not emails:
                st.warning("Nenhum email válido encontrado para envio.")
                return

            self.send_bulk_emails(emails, subject, body)

        except Exception as e:
            st.error(f"Erro no processamento: {str(e)}")
            logging.error(f"Erro no processamento: {str(e)}")

    def get_email_list(self, single_recipient, file, email_column) -> List[str]:
        """Obtém lista de emails para envio"""
        if single_recipient and EmailProcessor.validate_email(single_recipient):
            return [single_recipient]
        elif file and email_column:
            return EmailProcessor.process_excel(file, email_column)
        return []

    def send_bulk_emails(self, emails: List[str], subject: str, body: str):
        """Envia emails em massa com barra de progresso"""
        total = len(emails)
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        sender = EmailSender(self.config)
        successes = 0

        for idx, email in enumerate(emails):
            if sender.send_email(email, subject, body):
                successes += 1
            
            progress = (idx + 1) / total
            progress_bar.progress(progress)
            status_text.text(f"Processando: {idx + 1} de {total} emails")

        self.show_completion_message(successes, total)

    def show_completion_message(self, successes: int, total: int):
        """Exibe mensagem de conclusão"""
        if successes == total:
            st.success(f"Todos os {total} emails foram enviados com sucesso!")
        else:
            st.warning(f"Enviados {successes} de {total} emails. Alguns envios falharam.")

def main():
    ui = StreamlitUI()
    ui.sidebar()
    ui.main_page()

if __name__ == "__main__":
    main()