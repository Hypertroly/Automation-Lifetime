import win32com.client as win32
import datetime
import sys

outlook = win32.Dispatch('Outlook.Application')

sendfromAC=None
for oacc in outlook.Session.Accounts:
#    if oacc.SmtpAddress == "Movimentacoes Fundos - XP Investimentos <movimentacoes.fundos@xpi.com.br>":
    if oacc.SmtpAddress == 'Assessores (LIFETIME) <ASSESSORES@LIFETIMEINVEST.COM.BR>':
#    if oacc.SmtpAddress == 'Tiago Sousa <tiago.sousa@LIFETIMEASSET.COM.BR>':
        sendfromAC = oacc
        break

mail = outlook.CreateItem(0)

datahj = str(datetime.date.today())

datahj = datahj[8]+ datahj[9] + '/' + datahj[5]+datahj[6] + '/' + datahj[0:4]


if sendfromAC:
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, sendfromAC))
mail.To = 'Assessores (LIFETIME) <ASSESSORES@LIFETIMEINVEST.COM.BR>'
#mail.To = 'Tiago Sousa <tiago.sousa@LIFETIMEASSET.COM.BR>'
mail.Cc = 'Operations <operations@LIFETIMEASSET.COM.BR>'
mail.Subject = 'Acompanhamento diário Balanced'
mail.Attachments.Add(r"C:\Users\thiago.sousa\Desktop\Acompanhamento Fundos\Balan\Acompanhamento Diário Balanced.pdf")

attachment1 = mail.Attachments.Add(r'C:\movbtg\Assinatura.png')
attachment1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "Assinatura")

mail.HTMLBody = "<HTML lang='en' xmlns='http://www.w3.org/1999/xhtml' xmlns:o='urn:schemas-microsoft-com:office:office'> " \
                + "<head>" \
                + "<!--[if gte mso 9]><xml> \
                        <o:OfficeDocumentSettings> \
                        <o:Allowjpeg/> \
                        <o:PixelsPerInch>96</o:PixelsPerInch> \
                        </o:OfficeDocumentSettings> \
                    </xml> \
                    <![endif]-->" \
                + "</head>" \
                + "<BODY>"

emailbody = "<BR>Lifetime,<b> </b>" \
                + "<BR><BR> Segue o acompanhamento diário referente ao dia %s  </b> "% (datahj)\
                + "<BR><BR> Fundo: Lifetime Balanced FIC FIM </b><BR><BR><BR><BR><BR>"\
                + "<html><body><img src='cid:Assinatura'></body></html>"

mail.HTMLBody = mail.HTMLBody + emailbody

a=''
while a != 'y' and 'n':
    a=input('Email completo, deseja enviá-lo? (y/n) ')

    if a=='n':
        sys.exit()

mail.Send()
print('Email Balanced enviado')

sendfromAC=None
for oacc in outlook.Session.Accounts:
#    if oacc.SmtpAddress == "Movimentacoes Fundos - XP Investimentos <movimentacoes.fundos@xpi.com.br>":
    if oacc.SmtpAddress == 'Assessores (LIFETIME) <ASSESSORES@LIFETIMEINVEST.COM.BR>':
#    if oacc.SmtpAddress == 'Tiago Sousa <tiago.sousa@LIFETIMEASSET.COM.BR>':
        sendfromAC = oacc
        break

mail = outlook.CreateItem(0)

datahj = str(datetime.date.today())

datahj = datahj[8]+ datahj[9] + '/' + datahj[5]+datahj[6] + '/' + datahj[0:4]


if sendfromAC:
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, sendfromAC))
mail.To = 'Assessores (LIFETIME) <ASSESSORES@LIFETIMEINVEST.COM.BR>'
#mail.To = 'Tiago Sousa <tiago.sousa@LIFETIMEASSET.COM.BR>'
mail.Cc = 'Operations <operations@LIFETIMEASSET.COM.BR>'
mail.Subject = 'Acompanhamento diário Graphene'
mail.Attachments.Add(r"C:\Users\thiago.sousa\Desktop\Acompanhamento Fundos\Graph\Acompanhamento Diário Graphene.pdf")

attachment1 = mail.Attachments.Add(r'C:\movbtg\Assinatura.png')
attachment1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "Assinatura")

mail.HTMLBody = "<HTML lang='en' xmlns='http://www.w3.org/1999/xhtml' xmlns:o='urn:schemas-microsoft-com:office:office'> " \
                + "<head>" \
                + "<!--[if gte mso 9]><xml> \
                        <o:OfficeDocumentSettings> \
                        <o:Allowjpeg/> \
                        <o:PixelsPerInch>96</o:PixelsPerInch> \
                        </o:OfficeDocumentSettings> \
                    </xml> \
                    <![endif]-->" \
                + "</head>" \
                + "<BODY>"

emailbody = "<BR>Lifetime,<b> </b>" \
                + "<BR><BR> Segue o acompanhamento diário referente ao dia %s  </b> "% (datahj)\
                + "<BR><BR> Fundo: Lifetime Graphene FIA</b><BR><BR><BR><BR><BR>"\
                + "<html><body><img src='cid:Assinatura'></body></html>"

mail.HTMLBody = mail.HTMLBody + emailbody

a=''
while a != 'y' and 'n':
    a=input('Email completo, deseja enviá-lo? (y/n) ')

    if a=='n':
        sys.exit()

mail.Send()
print('Email Graphene enviado')