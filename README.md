# 📊 Consolidator de Movimentações Financeiras  
*Automação para integração e consolidação de dados por meio de APIs e e-mails*

---

## 🔎 Sobre o Projeto

Este projeto foi desenvolvido para **automatizar a coleta, leitura e consolidação de movimentações financeiras** utilizando duas fontes principais:

1. **APIs externas (ou internas)** que fornecem dados estruturados.
2. **E-mails recebidos no Outlook**, contendo planilhas ou arquivos de transações.

Ele atua como uma ferramenta de ETL (Extract, Transform, Load), reunindo informações dispersas em um único arquivo consolidado.

> 🔒 *O código disponibilizado é uma versão totalmente genérica e censurada, não contendo qualquer informação sensível, endpoint real ou regra corporativa específica.*

---

## 🚀 Funcionalidades

- 📥 **Leitura automática da caixa de entrada Outlook**
  - Filtragem por assuntos específicos
  - Validação de destinatários
  - Identificação e download automático de anexos

- 🔗 **Consulta a APIs**
  - Requisição HTTP GET
  - Tratamento de resposta JSON
  - Renomeação e padronização de colunas (genérica)

- 📁 **Tratamento e normalização dos dados**
  - Padronização de texto
  - Conversão de datas
  - Classificação de status e tipos de operação
  - Consolidação de múltiplas fontes

- 📊 **Geração de arquivo Excel consolidado**
  - Mescla dos dados das APIs + anexos
  - Salvamento automatizado com aviso se o arquivo estiver aberto

- ⚙️ **Automação do fluxo completo**
  - Buscar e-mails → extrair anexos → consultar APIs → consolidar dados → gerar Excel

---

## 🧱 Estrutura do Projeto

project/
│
├── consolidator.py # Script principal (exemplo censurado)
├── /conteudo/ # Pasta de saída (exemplo)
│ └── resultado_consolidado.xlsx
└── README.md # Documentação


---

## 📦 Requisitos

### 🔧 **Python 3.8+**

### 🧩 Bibliotecas utilizadas:

- pandas  
- requests  
- pywin32 (win32com.client)  
- tkinter  
- unicodedata  

Instale com:

```bash
pip install pandas requests pywin32
⚠️ O uso do Outlook requer Windows + Outlook instalado.

🛠 Instalação
Clone este repositório:

git clone https://github.com/LeonardoOGSilva/Consolidador.git
Instale as dependências:

pip install -r requirements.txt
Certifique-se de que:

O Outlook está instalado e configurado

As APIs de exemplo foram substituídas por URLs reais

Os assuntos e filtros foram ajustados para o seu ambiente

▶️ Como Usar
Execute o script:

python consolidator.py
O fluxo de execução será:

Conectar ao Outlook

Buscar e-mails com assuntos configurados

Baixar anexos para a pasta definida

Consultar APIs e carregar os dados

Consolidar informações

Gerar o arquivo final em Excel

Ao final, será exibida uma mensagem no console indicando que o processo foi concluído.

📁 Output
O script gera:

resultado_consolidado.xlsx
Este arquivo contém:

Dados vindos das APIs configuradas

Dados importados dos anexos recebidos via e-mail

Colunas padronizadas e consolidadas

⚠️ Limitações
Dependência do Microsoft Outlook (Windows)

Necessidade de acesso válido às APIs configuradas

Alguns comportamentos podem variar conforme configurações de segurança corporativa

Este projeto é uma versão genérica e sem regras reais de negócio

🛡️ Sobre Segurança
Esta versão do projeto foi completamente censurada e não contém:

URLs reais de API

Nomes de sistemas internos

Caminhos corporativos

Assuntos reais de e-mail

Estruturas sensíveis

Dados confidenciais

É segura para publicação pública.

🤝 Contribuições
Contribuições são bem-vindas!
Sinta-se livre para abrir Issues ou enviar Pull Requests.

📄 Licença
Este projeto pode ser utilizado, modificado e distribuído livremente, desde que mantida a estrutura de censura e sem incluir dados sensíveis reais.

📬 Contato
Caso tenha dúvidas ou queira adaptar o projeto ao seu ambiente real, entre em contato pelo GitHub issues.
