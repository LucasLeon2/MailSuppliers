import pandas as pd
from hardcore import NewMail, MailBody

newfile = './Planilha_fornecedores_email.xlsx'
dados = pd.read_excel(newfile)

#TODO CORRIGIR NOMES DAS VARIÁVEIS
codigo = dados['Código']
alvara = dados['Alvará']
contrato = dados['Contrato social']
demo = dados['Demonstrativo do Resultado']
lao = dados['LAO']
qaf = dados['QAF']
termo = dados['Responsabilidade Social']
iso9 = dados['ISO 9001']
iso14 = dados['ISO 14001']
iso45 = dados['ISO 45001']
email = dados['E-mail']

master = []
alv = []
con = []
dr = []
l = []
q = []
ter = []
i9 = []
i14 = []
i45 = []

print(dados)

for n in range(0, 100):

    alvarat = alvara.isnull()
    contratot = contrato.isnull()
    demot = demo.isnull()
    laot = lao.isnull()
    qaft = qaf.isnull()
    termot = termo.isnull()
    iso9t = iso9.isnull()
    iso14t = iso14.isnull()
    iso45t = iso45.isnull()

    if alvarat[n] == 1 or laot[n] == 1 or qaft[n] == 1 or termot[n] == 1:
        master.append(1)

    else:
        master.append(0)

    if alvarat[n] == 1:
        alv.append('Alvará de funcionamento')

    else:
        alv.append('')

    if laot[n] == 1:
        l.append('Licença ambiental de operação')

    else:
        l.append('')

    if qaft[n] == 1:
        q.append('Questionário de avaliação de fornecedor(em anexo para ser preenchido e enviado de volta)')

    else:
        q.append('')

    if termot[n] == 1:
        ter.append('Termo de responsabilidade social e ambiental(em anexo para ser assinado e enviado de volta)')

    else:
        ter.append('')

    if contratot[n] == 1:
        con.append('Contrato social')

    else:
        con.append('')

    if demot[n] == 1:
        dr.append('Demonstrativo de resultado de 2022')

    else:
        dr.append('')

    if iso9t[n] == 1:
        i9.append('ISO 9001')

    else:
        i9.append('')

    if iso14t[n] == 1:
        i14.append('ISO 14001')

    else:
        i14.append('')

    if iso45t[n] == 1:
        i45.append('ISO 45001')

    else:
        i45.append('')

for n in range(0, 383):

    if master[n] == 1:
        corpoemail = MailBody(alv[n], l[n], q[n], ter[n], con[n], dr[n], i9[n], i14[n], i45[n])
        novoemail = NewMail(corpoemail.text_block(), codigo[n], email[n])
        novoemail.new_mail()

