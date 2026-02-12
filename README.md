# CobrancaNF

Sistema automatizado para processamento de pendências de notas fiscais e envio de relatórios por e-mail, organizado por filial.

## Funcionalidades

- **Processamento de CSV**: Lê dados de exportação e gera planilhas Excel formatadas por filial.
- **Gestão de E-mails**: Interface gráfica para configurar e-mails de administradores e coordenadores por loja.
- **Automação de E-mail**: Integração com Outlook para envio automático dos relatórios.
- **Persistência**: Configurações salvas localmente em JSON (`Documents/CobrancaNF/email_config.json`).

## Estrutura do Projeto

```text
cobrancanf/
├── main.pyw                # Ponto de entrada (Interface Gráfica)
├── requirements.txt        # Dependências do projeto
├── build_exe.bat           # Script para gerar executável
└── src/
    ├── application/        # Lógica de negócio (Processamento de planilhas)
    ├── infrastructure/     # Serviços (E-mail, Gestão de Configuração)
    └── presentation/       # Interface Gráfica (Tkinter)
```

## Instalação

1. Certifique-se de ter o **Python 3.10+** instalado.
2. Clone o repositório ou baixe os arquivos.
3. Instale as dependências:

```bash
pip install -r requirements.txt
```

## Como Usar

### 1. Iniciar a Interface

Execute o arquivo `main.pyw` para abrir o gerenciador de e-mails:

```bash
python main.pyw
```

### 2. Configurar Filiais

- Use o botão **"Adicionar Filial"** para cadastrar uma nova loja.
- Adicione os e-mails dos responsáveis nas seções correspondentes.
- O sistema criará automaticamente o arquivo de configuração no primeiro uso.

### 3. Processar Pendências

A interface permite carregar o arquivo CSV exportado do sistema para processamento e envio automático via Outlook.

### Exemplo do arquivo CSV

```csv
Loja;CNPJ;Fornecedor;Nr. IE;Dt. Emissão;Mod.;Série;Nr. Documento;Vlr. Documento;Evento;Observações
2;00.000.000/0000-00;Empresa Ltda;'000000000';18/02/2025;55;1;000000000;000,00;CIÊNCIA DA OPERAÇÃO;
```

## Gerar Executável (.exe)

Para distribuir o sistema como um aplicativo Windows:

1. Execute o script `build_exe.bat`.
2. O executável será gerado na pasta `dist/CobrancaNF.exe`.

---

**Observação**: O envio de e-mails requer o **Microsoft Outlook** instalado e configurado na máquina.
