# Automação XML MoneyCarWeb

📌 Sobre o Projeto
Este script automatiza o login e o download de arquivos XML do MoneyCarWeb. Ele usa Selenium para navegar no site, baixar os arquivos e salvar os dados em uma planilha Excel.

🛠️ Tecnologias Utilizadas
Python
Selenium
WebDriver Manager
openpyxl
JSON
Logging
⚙️ Configuração do Ambiente
Requisitos
Antes de começar, você precisa:
✔ Ter o Python 3 instalado
✔ Ter o Google Chrome instalado
✔ Usar o WebDriver Manager para gerenciar o ChromeDriver

Instalando as Dependências
Basta rodar:

bash
Copiar
Editar
pip install selenium webdriver-manager openpyxl
🚀 Como Usar
1️⃣ Configurar as Credenciais
Antes de rodar o script, edite o código e defina suas credenciais:

python
Copiar
Editar
USER = "seu_email"
SENHA = "sua_senha"
DOWNLOAD_DIR = "C:\\Users\\seu_usuario\\Downloads\\XML"
2️⃣ Executar o Script
Para iniciar a automação, rode:

bash
Copiar
Editar
python automacao_xml.py
3️⃣ O que o script faz?
🔹 Acessa o MoneyCarWeb e faz login automaticamente.
🔹 Navega até a página de faturamento.
🔹 Seleciona o período desejado.
🔹 Baixa os arquivos XML.
🔹 Salva o progresso para evitar downloads repetidos.

Se algo der errado, os erros serão registrados no arquivo automacao_xml.log.

❗ Possíveis Problemas e Soluções
🔹 "Erro ao acessar página de XMLs"
✔ Verifique se o MoneyCarWeb está online.
✔ Confirme suas credenciais.

🔹 "Timeout esperando download do XML"
✔ Verifique sua conexão com a internet.
✔ Ajuste o tempo de espera no código, caso necessário.

🤝 Contribuição
Quer contribuir? Siga estes passos:

Faça um fork do repositório.
Crie uma branch (git checkout -b minha-feature).
Faça suas alterações e commit (git commit -m "Minha nova feature").
Envie um push (git push origin minha-feature).
Abra um Pull Request.
📜 Licença
Este projeto é distribuído sob a licença MIT.

