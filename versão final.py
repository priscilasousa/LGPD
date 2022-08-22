import win32com.client as win32                     # biblioteca para integração com Outlook
from pathlib import Path                            # biblioteca para caminho do arquivo
from fpdf import FPDF                               # biblioteca para ler pdf
import re                                           # biblioteca para usar regex
import PyPDF2                                       # biblioteca para
import os                                           # biblioteca para
import pandas as pd                                 # biblioteca para manipular dados
import docx2txt                                     # biblioteca para manipular .docx
import shutil

# criando uma pasta para salvar todos os arquivos
destino = Path.cwd() / "output"
destino.mkdir(parents=True, exist_ok=True)
os.chdir(r"C:\Users\Priscila\PycharmProjects\LGPD\output")

# criando conexão com o Outlook
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Acessando a pasta de enviados
sent = outlook.GetDefaultFolder(5)

# Pegando ultimo item enviado da pasta Sent
messages = sent.items
message = messages.GetLast()

# Salvando todos os itens da pasta
subject = message.subject                                       # pega subject
body = message.body                                             # pega corpo da mensagem
attachments = message.Attachments                               # pega anexo

name = str(subject).replace(':', "").replace('/', "") + '.msg'  # formata nome do email
message.SaveAs(os.getcwd() + '//' + name)                       # salva email na pasta

# Criando pasta com o mesmo subject do email
pasta_destino = destino / str(subject).replace(':',"").replace('/',"")
print(pasta_destino)
pasta_destino.mkdir(parents=True, exist_ok=True)

# criando arquivo com o corpo do email
Path(pasta_destino / 'Corpo_email.txt').write_text(str(body))

# Transformando o corpo do email em PDF
pdf = FPDF()
pdf.add_page()
pdf.set_font("Arial", size=10)
subject2= subject.replace(':',"").replace('/',"")
f = open(f'C:/Users/Priscila/PycharmProjects/LGPD/output/{subject2}/Corpo_email.txt', "r")
for x in f:
    pdf.cell(200, 10, txt=x, ln=1, align='C')
pdf.output(f"C:/Users/Priscila/PycharmProjects/LGPD/output/{subject2}/Corpo_email.pdf")

# Lendo corpo do email em PDF
pdf_corpo = open(f"C:/Users/Priscila/PycharmProjects/LGPD/output/{subject2}/Corpo_email.pdf",'rb')
dados_corpo = PyPDF2.PdfFileReader(pdf_corpo)
numero_pag_corpo=dados_corpo.numPages           # Contando o número de páginas do arquivo

if numero_pag_corpo == 1:                       # Lendo arquivo com 1 página
    pag_corpo = dados_corpo.getPage(0)
    txt_corpo = pag_corpo.extractText()
    txt_corpo = ''.join(txt_corpo).lower()
    txt_corpo = re.sub('\n', '', txt_corpo)
else:
    i = 1
    while (i <= numero_pag_corpo):                  # Lendo arquivo com mais de uma página
        pag_corpo = dados_corpo.getPage(i - 1)
        txt_corpo = pag_corpo.extractText()
        txt_corpo = ''.join(txt_corpo)
        txt_corpo = re.sub('\n', '', txt_corpo)
        i = i + 1

# Salvando anexos na pasta (caso exista mais de um)
for att in attachments:
    att.SaveAsFile(pasta_destino / str(att))
    file_extension = Path(rf'C:\Users\Priscila\PycharmProjects\LGPD\output\{subject}\{att}').suffix

    if file_extension == '.xlsx':
        # Lendo os anexos do email em XLSX
        xlsx = pd.read_excel(rf'C:\Users\Priscila\PycharmProjects\LGPD\output\{subject}\{att}')
        xlsx = xlsx.to_string()
        csv = 'zero'
        txt_anexo = 'zero'
    if file_extension == '.pdf':
        # Lendo os anexos do email em PDF
        pdf_file = open(f"C:/Users/Priscila/PycharmProjects/LGPD/output/{subject2}/{att}", 'rb')
        dados_anexo = PyPDF2.PdfFileReader(pdf_file)
        numero_pag_anexo = dados_anexo.numPages

        if numero_pag_anexo == 1:
            pag_anexo = dados_anexo.getPage(0)
            txt_anexo = pag_anexo.extractText()
            txt_anexo = ''.join(txt_anexo).lower()
            txt_anexo = re.sub('\n', '', txt_anexo)
            # print('1 pagina')
        else:
            i = 1
            while (i <= numero_pag_anexo):
                pag_anexo = dados_anexo.getPage(i - 1)
                txt_anexo = pag_anexo.extractText()
                txt_anexo = ''.join(txt_anexo)
                txt_anexo = re.sub('\n', '', txt_anexo)
                i = i + 1
                # print('mais de uma página')
    if file_extension == '.csv':
        csv = pd.read_csv(rf'C:\Users\Priscila\PycharmProjects\LGPD\output\{subject}\{att}', sep='\t', encoding='latin-1', header=None)
        csv = csv.to_string()
    if file_extension == '.docx':
        doc = docx2txt.process(rf'C:\Users\Priscila\PycharmProjects\LGPD\output\{subject}\{att}')
    if file_extension == '.txt':
        txt = open(rf'C:\Users\Priscila\PycharmProjects\LGPD\output\{subject}\{att}')
        txt = txt.read().lower()

# Se não tiver anexo, txt_anexo é igual a 'zero'
if (len(attachments)) == 0 or file_extension != '.pdf':
    txt_anexo = 'zero'
if file_extension != '.xlsx':
    xlsx = 'zero'
if file_extension != '.csv':
    csv = 'zero'
if file_extension != '.docx':
    doc = 'zero'
if file_extension != '.txt':
    txt = 'zero'

# Criando email para envio
outlook2 = win32.Dispatch('outlook.application')
email = outlook2.CreateItem(0)

# Configurando as informações do email
email.To = "priscilla.sooousa@gmail.com"
email.Subject = "[ALERTA] Possível vazamento de Informações"
email.HTMLBody ="""
<p>Caro administrador de Segurança,</p>

<p>Este usuário está tentando enviar um email que pode conter informações sensíveis.</p>

<p>Por favor, verificar e tomar as ações devidas, caso necessário!</p>
<p>Este é um email automático, por favor, não responder!</p>
"""

email.Attachments.Add(rf'C:\Users\Priscila\PycharmProjects\LGPD\output\{name}')

# Regex
masc_r = re.compile('(católico)|(crente)|(umbandista)|(protestante)|(espírita)|(candomblecista)|(candomblécista)|(espirita)|(catolico)')
masc_o = re.compile('((orientação) (sexual))|(gay)|(homossexual)|(lésbica)|(bissexual)|(heterossexual)|(transsexual)|(travesti)|(traveco)')
masc_c = re.compile('(^(\d{5})-(\d{3})$)|(\d{8})')
masc_cpfnpj = re.compile('([0-9]{2}[\.]?[0-9]{3}[\.]?[0-9]{3}[\/]?[0-9]{4}[-]?[0-9]{2})|([0-9]{3}[\.]?[0-9]{3}[\.]?[0-9]{3}[-]?[0-9]{2})')

# Validação
if re.search(masc_r, txt_corpo) or re.search(masc_r, xlsx) or re.search(masc_r, csv) or re.search(masc_r, txt_anexo) or re.search(masc_r, doc) or re.search(masc_r, txt):
    email.Send()
    #print("Intolerancia religiosa!")
elif re.search(masc_cpfnpj, txt_corpo) or re.search(masc_cpfnpj, xlsx) or re.search(masc_cpfnpj, csv) or re.search(masc_cpfnpj, txt_anexo) or re.search(masc_cpfnpj, doc) or re.search(masc_cpfnpj, txt):
    email.Send()
    #print("CNPJ ou CPF!")
elif re.search(masc_c, txt_corpo) or re.search(masc_c, xlsx) or re.search(masc_c, csv) or re.search(masc_c, txt_anexo) or re.search(masc_c, doc) or re.search(masc_c, txt):
    email.Send()
    #print('CEP!')
elif re.search(masc_o, txt_corpo) or re.search(masc_o, xlsx) or re.search(masc_o, csv) or re.search(masc_o, txt_anexo) or re.search(masc_o, doc) or re.search(masc_o, txt):
    email.Send()
    #print('Sexualidade!')
else:
    print('Não encontrado')

#filePath=Path(rf'C:\Users\Priscila\PycharmProjects\LGPD\output\{subject}')
#filePath.unlink()
#shutil.rmtree(rf'C:\Users\Priscila\PycharmProjects\LGPD\output')
