# Protect/UnProtect Excel Worksheets
Criado para remover a senha de planilhas diretamente pelo arquivo XML. Como apenas a SHA-512 da senha é armazenada, um arquivo .JSON é gerado para que possa-se proteger a senha novamente.
# Tecnologias empregadas
  <ul>
    <li>jsZIP para acessar os arquivos .xml das planilhas</li>
    <li>xml-js para ler e modificar os arquivos .xml</li>
    <li>electron para interface de usuário</li>
  </ul>

# Instalação
Para instalar as depêndencias:<br>
<code>npm i</code><br>
Para gerar o instalador usar os seguintes comandos:<br>
<code>npm run build</code> 

