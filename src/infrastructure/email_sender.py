from datetime import datetime
import win32com.client as win32
import locale
from src.infrastructure.config_manager import EmailConfigManager
from src.infrastructure.config.settings import Settings

class SendEmail:
    """Gerenciador de envio de e-mails via Outlook"""

    def __init__(self):
        # Inicializa o gerenciador de configuração
        self._config_service = EmailConfigManager()
        self._settings = Settings()
        try:
            locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")  # no Windows, pode ser "Portuguese_Brazil.1252"
        except locale.Error:
            try:
                locale.setlocale(locale.LC_TIME, "Portuguese_Brazil.1252")
            except locale.Error:
                pass # Fallback to default locale if both fail

    def execute(self, loja: str, periodoInicial: datetime, periodoFinal: datetime, destinatario: list[str], copia: list[str], caminho_arquivo: str, qt, vlrTotal, table):
        """Envia o e-mail com a planilha de pendências em anexo"""
        outlook = win32.Dispatch('Outlook.Application')
        email = outlook.CreateItem(0)

        company = self._settings.COMPANY

        email.Display()
        assinatura_completa = email.HTMLBody

        email.Subject = f"PENDÊNCIA DE LANÇAMENTO - LOJA {loja} - {periodoFinal.strftime('%m/%Y')}"
        email.HTMLBody = f"""
        <html>
            <body>
                <div style="font-family: Aptos; font-size: 12pt;">
                <p>{self._saudacao()}</p>
                <p>
                Segue em anexo a planilha com as notas fiscais e CTEs que ainda estão pendentes 
                de lançamento no sistema, {self._check_period(periodoInicial, periodoFinal)}, e que não possuem 
                justificativa registrada.
                </p>

                <p>Pedimos, por favor, que:</p>

                <ul>
                <li>Lancem no sistema as notas que ainda não foram lançadas;</li>
                <li>Verifiquem cada pendência com atenção;</li>
                <li>Confiram se a mercadoria realmente não chegou, se não será mais entregue ou se é necessário fazer a recusa no portal GED;</li>
                <li>Registrem a justificativa no sistema, no campo <strong>“OBSERVAÇÕES”</strong>, pela filial.</li>
                </ul>

                <p>
                É muito importante que a informação esteja registrada no sistema, 
                pois precisamos informar ao time da {company} o motivo pelo qual essas notas ainda não deram entrada.
                </p>
                <p style="font-family: Aptos; font-size: 14pt;">Quantidade de documentos: {qt}</p>
                <p style="font-family: Aptos; font-size: 14pt;">Valor total: R${vlrTotal}</p><br/>
                {table}
                {assinatura_completa}
                </div>
            </body>
        </html>
        """
        email.Attachments.Add(caminho_arquivo)

        email.To = ";".join(destinatario) if len(destinatario) > 0 else ""
        email.CC = ";".join(copia) if len(copia) > 0 else ""
        email.Send()

    def _saudacao(self):
        """Retorna saudação apropriada baseada na hora atual"""
        hora_atual = datetime.now().hour

        if hora_atual < 12:
            return "Bom dia!"
        elif hora_atual < 18:
            return "Boa tarde!"
        else:
            return "Boa noite!"

    def _check_period(self, periodoInicial: datetime, periodoFinal: datetime):
        """Retorna a string formatada do período"""
        # diferença total em meses
        diff_meses = (periodoFinal.year - periodoInicial.year) * 12 + (periodoFinal.month - periodoInicial.month)

        # ajusta se o dia final ainda não completou o mês
        if periodoFinal.day < periodoInicial.day:
            diff_meses -= 1

        if diff_meses >= 1:
            return f"no período entre {periodoInicial.strftime('%m/%Y')} até {periodoFinal.strftime('%m/%Y')}"
        else:
            return f"no mês {periodoInicial.strftime('%m/%Y')}"

    def get_admins(self, filial: str) -> list[str]:
        """Retorna lista de e-mails dos administradores da filial a partir do arquivo de configuração"""
        return self._config_service.get_store_admins(filial)

    def get_coordinators(self, filial: str) -> list[str]:
        """Retorna lista de e-mails dos coordenadores da filial a partir do arquivo de configuração"""
        return self._config_service.get_store_coordinators(filial)

    def get_test_emails(self, filial: str):
        """Retorna lista de e-mails de teste"""
        email_test = self._settings.EMAIL_TEST
        return [email_test] if email_test else []