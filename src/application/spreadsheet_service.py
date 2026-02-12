import pandas as pd
import os
from openpyxl import load_workbook
from src.infrastructure.email_sender import SendEmail
from src.infrastructure.config_manager import EmailConfigManager
from src.domain.entities import BranchReport

class SpreadsheetService:
    """Serviço para processamento de planilhas e orquestração de envio de e-mails"""
    
    def __init__(self):
        self._sendemail = SendEmail()
        self._config_manager = EmailConfigManager()
    
    def execute(self, csv_path, output_base):
        """Lê o CSV, gera planilhas por filial e envia e-mails"""
        # Lê o CSV
        df = pd.read_csv(csv_path, sep=";", encoding="latin1", dtype=str)

        df.columns = df.columns.str.strip()

        # Remove coluna Observações se existir
        if "Observações" in df.columns:
            df = df.drop(columns=["Observações"])

        # Converte Vlr. Documento para float
        col_valor = "Vlr. Documento"
        if col_valor in df.columns:
            # Substitui '.' por vazio e ',' por '.' para converter em float
            df[col_valor] = df[col_valor].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
            # Arredonda para 2 casas decimais
            df[col_valor] = df[col_valor].round(2)

        df['Dt. Emissão'] = pd.to_datetime(df['Dt. Emissão'], dayfirst=True, errors='coerce')

        periodoInicial = df['Dt. Emissão'].min()
        periodoFinal = df['Dt. Emissão'].max()

        # Verifica o nome exato da coluna da filial
        col_filial = "Loja"
        if col_filial in df.columns:
            df[col_filial] = df[col_filial].apply(lambda x: f"0{x}" if len(str(int(x))) < 2 else str(x))

        # Garante que a pasta de saída exista
        os.makedirs(output_base, exist_ok=True)

        reports = []
        # Gera um arquivo por filial
        for filial, grupo in df.groupby(col_filial):
            filial_str = str(filial)
            pasta_filial = os.path.join(output_base, filial_str)
            pasta_periodo = os.path.join(pasta_filial, periodoFinal.strftime("%m%Y"))
            os.makedirs(pasta_filial, exist_ok=True)
            os.makedirs(pasta_periodo, exist_ok=True)
            arquivo_excel = os.path.join(pasta_periodo, f"Pendencias {filial_str}.xlsx")
            
            # Salva o arquivo Excel
            grupo.to_excel(arquivo_excel, index=False, engine="openpyxl")
            
            soma_valor = grupo['Vlr. Documento'].sum()
            soma_formatada = f"{soma_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

            grupo['Ano-Mes'] = grupo['Dt. Emissão'].dt.to_period('M')
            contagem_por_mes = grupo.groupby('Ano-Mes').size().reset_index(name='Quantidade por mês')
            
            # Gera tabela HTML para o corpo do e-mail
            html_table = contagem_por_mes.to_html(
                index=False,
                border=0,
                justify="center",
                classes="table",
                table_id="tabela_mes"
            )
            grupo = grupo.drop(columns=["Ano-Mes"])
            
            # Autoajusta colunas no Excel
            wb = load_workbook(arquivo_excel)
            ws = wb.active

            for col in ws.columns:
                max_length = 0
                col_letter = col[0].column_letter
                for cell in col:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = max_length + 2
            
            wb.save(arquivo_excel)
            
            report = BranchReport(
                branch=filial_str,
                period_initial=periodoInicial,
                period_final=periodoFinal,
                excel_path=arquivo_excel,
                quantity=len(grupo),
                total_value=f"{soma_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                table=html_table
            )

            reports.append(report)
        
        return reports