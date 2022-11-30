import pandas as pd
import win32com.client as win32

outlook = win32.Dispatch("Outlook.Application")

anexoqaf = "C:\\Users\lucas.leon\Desktop\Envio_for\F-077 - 12 - QAF_SEQ - REV.12.xlsx" #TODO: ATUALIZAR CAMINHO
anexotermo = "C:\\Users\lucas.leon\Desktop\Envio_for\SOCIAL AND ENVIRONMENTAL RESPONSIBILITY COMMITMENT v2.pdf" #TODO: ATUALIZAR CAMINHO


class SaveMail:

    def __init__(self, mail_body, codigo, email):
        self.mail = outlook.CreateItem(0)

        self.body = mail_body

        self.codigo = codigo

        self.email = email

        self.subject = f'Documentação - Docol - {self.codigo}'

    def new_mail(self):
        self.mail.Subject = self.subject
        self.mail.To = self.email
        #self.mail.HTMLBody = self.body #Open the window with email text
        self.mail.Display()

        if qaft[n] == 1:
            self.mail.Attachments.Add(anexoqaf)

        if termot[n] == 1:
            self.mail.Attachments.Add(anexotermo)

        index = self.mail.HTMLbody.find('>', self.mail.HTMLbody.find('<body'))
        self.mail.HTMLbody = self.mail.HTMLbody[:index + 1] + self.body + self.mail.HTMLbody[index + 1:]
        self.mail.Save()
        self.mail.Send()


class MailBody:

    def __init__(self, alv, l, q, ter, con, dr, i9, i14, i45): #TODO CORRIGIR NOMES DAS VARIÁVEIS
        self.alv = alv
        self.l = l
        self.q = q
        self.ter = ter
        self.con = con
        self.dr = dr
        self.i9 = i9
        self.i14 = i14
        self.i45 = i45
        self.text = ""

    def text_block(self):

        alv = self.alv
        l = self.l
        q = self.q
        ter = self.ter
        con = self.con
        dr = self.dr
        i9 = self.i9
        i14 = self.i14
        i45 = self.i45

        self.text = """
            <body style="background-color: #FFFFFF; margin: 0; padding: 0; -webkit-text-size-adjust: none; text-size-adjust: none;">
                <table class="nl-container" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #FFFFFF;">
                   <tbody>
                            <tr>
                                <td>
                                    <table class="row row-1" align="center" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                        <tbody>
                                            <tr>
                                                <td>
                                                    <table class="row-content stack" align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; color: #000000; width: 500px;" width="500">
                                                        <tbody>
                                                            <tr>
                                                                <td class="column column-1" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; padding-top: 5px; padding-bottom: 5px; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;">
                                                                    <table class="paragraph_block block-1" width="100%" border="0" cellpadding="10" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                                                                        <tr>
                                                                            <td class="pad">
                                                                                <div style="color:#000000;direction:ltr;font-family:Arial, Helvetica Neue, Helvetica, sans-serif;font-size:13px;font-weight:400;letter-spacing:0px;line-height:120%;text-align:left;mso-line-height-alt:15.6px;">
                                                                                    <p style="margin: 0; margin-bottom: 16px;">Caro fornecedor,</p>
                                                                                    <p style="margin: 0; margin-bottom: 16px;">A Docol demanda alguns documentos de seus fornecedores e alguns dos que temos da sua empresa estão desatualizados.</p>
                                                                                    <p style="margin: 0; margin-bottom: 16px;">Alguns documentos são imprescindíveis e outros opcionais. É interessante que enviem os opcionais também, caso possuam.</p>
                                                                                    <p style="margin: 0; margin-bottom: 16px;">Abaixo segue uma lista dos documentos&nbsp;<strong>imprescindíveis</strong> que estão desatualizados, seguida da lista dos documentos opcionais.(Pontos sem texto são esperados e referentes a documentos que não demandam atualização)</p>
                                                                                    <p style="margin: 0;">Documentos <strong>imprescindíveis</strong>:</p>
                                                                                </div>
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                    <table class="list_block block-2" width="100%" border="0" cellpadding="10" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                                                                        <tr>
                                                                            <td class="pad">
                                                                                <ul start="1" style="margin: 0; padding: 0; margin-left: 20px; list-style-type: revert; color: #000000; direction: ltr; font-family: Arial, Helvetica Neue, Helvetica, sans-serif; font-size: 13px; font-weight: 400; letter-spacing: 0px; line-height: 120%; text-align: left;">
                                                                                    {}
                                                                                    {}
                                                                                    {}
                                                                                    {}
                                                                                </ul>
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                    <table class="paragraph_block block-3" width="100%" border="0" cellpadding="10" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                                                                        <tr>
                                                                            <td class="pad">
                                                                                <div style="color:#000000;direction:ltr;font-family:Arial, Helvetica Neue, Helvetica, sans-serif;font-size:13px;font-weight:400;letter-spacing:0px;line-height:120%;text-align:left;mso-line-height-alt:15.6px;">
                                                                                    <p style="margin: 0;">Documentos não imprescindíveis:</p>
                                                                                </div>
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                    <table class="list_block block-4" width="100%" border="0" cellpadding="10" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                                                                        <tr>
                                                                            <td class="pad">
                                                                                <ul start="1" style="margin: 0; padding: 0; margin-left: 20px; list-style-type: revert; color: #000000; direction: ltr; font-family: Arial, Helvetica Neue, Helvetica, sans-serif; font-size: 13px; font-weight: 400; letter-spacing: 0px; line-height: 120%; text-align: left;">
                                                                                    {}
                                                                                    {}
                                                                                    {}
                                                                                    {}
                                                                                    {}
                                                                                </ul>
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                    <table class="paragraph_block block-5" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                                                                        <tr>
                                                                            <td class="pad">
                                                                                <div style="color:#101112;direction:ltr;font-family:Arial, Helvetica Neue, Helvetica, sans-serif;font-size:13px;font-weight:400;letter-spacing:0px;line-height:120%;text-align:left;mso-line-height-alt:15.6px;">
                                                                                    <p style="margin: 0; margin-bottom: 16px;">É importante frisar o caráter urgente dessa demanda. Sendo assim, peço que esse e-mail seja respondido com a documentação solicitada o quanto antes.</p>
                                                                                    <p style="margin: 0; margin-bottom: 16px;">Agradeço a atenção e conto a colaboração da sua empresa.</p>
                                                                                    <p style="margin: 0;">Att.,</p>
                                                                                </div>
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                </td>
                                                            </tr>
                                                        </tbody>
                                                    </table>
                                                </td>
                                            </tr>
                                        </tbody>
                                    </table>
                                    <table class="row row-2" align="center" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                        <tbody>
                                            <tr>
                                                <td>
                                                    <table class="row-content stack" align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; color: #000000; width: 500px;" width="500">
                                                        <tbody>
                                                            <tr>
                                                                <td class="column column-1" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; padding-top: 5px; padding-bottom: 5px; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;">
                                                                    <table class="icons_block block-1" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                                        <tr>
                                                                            <td class="pad" style="vertical-align: middle; color: #9d9d9d; font-family: inherit; font-size: 15px; padding-bottom: 5px; padding-top: 5px; text-align: center;">
                                                                                <table width="100%" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                                                                    <tr>
                                                                                        <td class="alignment" style="vertical-align: middle; text-align: center;">
                                                                                            <!--[if vml]><table align="left" cellpadding="0" cellspacing="0" role="presentation" style="display:inline-block;padding-left:0px;padding-right:0px;mso-table-lspace: 0pt;mso-table-rspace: 0pt;"><![endif]-->
                                                                                            <!--[if !vml]><!-->
                                                                                            <table class="icons-inner" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; display: inline-block; margin-right: -4px; padding-left: 0px; padding-right: 0px;" cellpadding="0" cellspacing="0" role="presentation">
                                                                                                <!--<![endif]-->
                                                                                            </table>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                </td>
                                                            </tr>
                                                        </tbody>
                                                    </table>
                                                </td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                """

        ladoA = '<li style="margin-bottom: 0px;">'
        ladoB = '</li>'

        if alvarat[n] == 1:
            alv = ladoA + alv + ladoB

        if laot[n] == 1:
            l = ladoA + l + ladoB

        if qaft[n] == 1:
            q = ladoA + q + ladoB

        if termot[n] == 1:
            ter = ladoA + ter + ladoB

        if contratot[n] == 1:
            con = ladoA + con + ladoB

        if demot[n] == 1:
            dr = ladoA + dr + ladoB

        if iso9t[n] == 1:
            i9 = ladoA + i9 + ladoB

        if iso14t[n] == 1:
            i14 = ladoA + i14 + ladoB

        if iso45t[n] == 1:
            i45 = ladoA + i45 + ladoB

        return self.text.format(alv, l, q, ter, con, dr, i9, i14, i45)

dados = pd.read_excel(r'C:\\Users\lucas.leon\Desktop\Envio_for\Planilha_fornecedores_email.xlsx') #TODO CORRIGIR CAMINHO

#TODO CORRIGIR NOMES DAS VARIÁVEIS
codigo = dados['Código']
alvara = dados['Alvará']
contrato = dados['Contrato social']
demo = dados['Demonstrativo do resultado']
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
        novoemail = SaveMail(corpoemail.text_block(), codigo[n], email[n])
        novoemail.new_mail()

