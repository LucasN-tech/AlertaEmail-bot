

<h1>Alerta de E-mail<h1>

### O Alerta de E-mail é um software desenvolvido em C# que realiza verificação de documentos e envia e-mails com base em certas condições. Ele foi projetado para automatizar o processo de alerta por e-mail, facilitando a comunicação com clientes.



# Funcionalidades

<li>Leitura de arquivo Excel: O software lê um arquivo de planilha do Excel que contém informações dos clientes, como CNPJ, razão social, e-mails, entre outros.</li>
<li>Verificação de documentos: Utilizando os dados dos clientes, o software realiza uma análise de documentos específicos associados a cada cliente.
Envio de e-mails: Com base nos resultados da verificação, o software envia e-mails para os clientes, seguindo lógicas distintas para diferentes cenários.</li>
<li>Gerenciamento de exceções: Em caso de erros durante o processo, o software captura as exceções e envia e-mails de exceção para um endereço de e-mail específico.
</li>



# Dependências
 O software utiliza as seguintes bibliotecas:

<li>System: Fornece tipos e funcionalidades fundamentais do .NET Framework.</li>
<li>System.Data: Fornece acesso a dados e manipulação de bancos de dados.</li>
<li>System.Net.Mail: Permite o envio de e-mails através de um servidor SMTP.</li>
<li>DocumentFormat.OpenXml.Bibliography: Biblioteca para trabalhar com referências bibliográficas no formato Open XML.</li>
<li>DocumentFormat.OpenXml.Wordprocessing: Biblioteca para trabalhar com documentos no formato Open XML.</li>
<li>ClosedXML.Excel: Biblioteca para trabalhar com planilhas do Excel.</li>


# Configuração e Uso

Para utilizar o Alerta de E-mail, siga estas etapas:

<li>Certifique-se de ter as dependências mencionadas instaladas.
Defina o caminho do arquivo de planilha Excel no código-fonte (var filePath).</li>

<li>Configure as informações do servidor SMTP do Outlook, incluindo o servidor, endereço de e-mail e senha (var smtpServer, var email, var password).</li>

<li>Execute o programa e aguarde a conclusão do processo de verificação e envio de e-mails.</li>

# Observações

<li>Certifique-se de que o arquivo de planilha Excel esteja formatado corretamente, seguindo a estrutura esperada pelo software.</li>
<li>Verifique se o servidor SMTP do Outlook está corretamente configurado e permitindo o envio de e-mails.
</li>

Este software é uma ferramenta útil para automatizar o 
processo de alerta por e-mail com base em condições específicas. Ele pode ser adaptado e personalizado de acordo com as necessidades do seu projeto.