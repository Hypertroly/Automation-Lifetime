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

datahj = datetime.date.today()
print(datahj)
dia=str(datahj)
dia=int(dia[8] + dia[9])
print(datahj.weekday())
if datahj.weekday()==0:
    dia=dia-3
elif datahj.weekday()==1:
    dia=dia-1
else:
    dia = str(dia-1)

dia=int(dia)
datahj = str(datahj)
diafsem = str(abs(dia-3))

diames1 = str(abs(dia-31))
diames2 = str(abs(dia-30))
diames3 = str(abs(dia-29))
dia = str(dia)

if dia=='1' or dia=='2' or dia=='3' or dia=='4' or dia=='5' or dia=='6' or dia=='7' or dia=='8' or dia=='9':
    dia='0'+ dia

if diafsem=='1' or diafsem=='2' or diafsem=='3' or diafsem=='4' or diafsem=='5' or diafsem=='6' or diafsem=='7' or diafsem=='8' or diafsem=='9':
    diafsem='0'+ diafsem

mes=datahj[5]+datahj[6]

mesp = '0'+str(int(mes)-1)

datasem = diafsem + '/' + mes + '/' + datahj[0:4]

diames1 = diames1 + '/' + mesp + '/' + datahj[0:4]
diames2 = diames2 + '/' + mesp + '/' + datahj[0:4]
diames3 = diames3 + '/' + mesp + '/' + datahj[0:4]

datahj = dia + '/' + mes + '/' + datahj[0:4]

cominterna = 'Operations <operations@LIFETIMEASSET.COM.BR>; Aline Paixão <aline.paixao@lftm.com.br>; Alyne Lins <alyne.lins@lftm.com.br>; Amanda Suosa Aita <amanda.aita@lftm.com.br>; Ana Paula Lainetti <ana.lainetti@lftm.com.br>; Andres Escobar <andres.escobar@LIFETIMEASSET.COM.BR>; Beatriz Santos <beatriz.santos@lftm.com.br>; Caio Roschel - Lifetime Invest <caio.roschel@lifetimeinvest.com.br>; Carlos Schincariol <carlos.schincariol@lftm.com.br>; Carlos Schincariol - Lifetime Invest <carlos.schincariol@lifetimeinvest.com.br>; CHARLIE <CHARLIE@lftm.com.br>; diretoria <diretoria@lftm.com.br>; Eduardo Rímoli <eduardo.rimoli@lftm.com.br>; Fernanda Moura <fernanda.moura@lftm.com.br>; Fernando Katsonis <fernando@lftm.com.br>; Fernando Katsonis (LFTMA) <fernando@LIFETIMEASSET.COM.BR>; Fernando Morales <fernando.morales@lftm.com.br>; Gabriel Brandão - Lifetime Invest <gabriel.brandao@lifetimeinvest.com.br>; Giovanni Marazzi <giovanni.marazzi@lftm.com.br>; Gisela Luti <gisela.luti@lftm.com.br>; Guilherme Burger <guilherme@lftm.com.br>; Josian Teixeira <josian.teixeira@LIFETIMEASSET.COM.BR>; Kevin Freundt <kevin.freundt@LIFETIMEASSET.COM.BR>; Laura Dantas Marques <laura.marques@lftm.com.br>; Liliana Praconi Berthier Jacomino <liliana.jacomino@lftm.com.br>; Luciane Torres <luciane.torres@lftm.com.br>; Luis Almeida <luis.almeida@lftm.com.br>; Marcello Popoff <marcello@lftm.com.br>; Mayara Fernanda Castro Silva Prieto <mayara.prieto@lftm.com.br>; Pedro Chaves <pedro.chaves@lftm.com.br>; Rafael Canto <rafael.canto@lftm.com.br>; Rafael Inaimo Chow <rafael.chow@LIFETIMEASSET.COM.BR>; Renato Klajner <renato.klajner@LIFETIMEASSET.COM.BR>; Renne Xavier <renne.xavier@lftm.com.br>; Rodolfo Rosina <rodolfo.rosina@lftm.com.br>; Suellen Silva <suellen.silva@lftm.com.br>; Thais Neves <thais.neves@lftm.com.br>; Tiago Nunes Bernardes de Sousa <tiago.sousa@LIFETIMEASSET.COM.BR>; vida <vida@lftm.com.br>; Vinicius Rocha <vinicius.rocha@LIFETIMEASSET.COM.BR>; Vitor Carettoni <vitor.carettoni@lftm.com.br>; Vitor Fernandes Carettoni <vitor.A65681@agenteinvest.com.br>; Wallace Nascimento <wallace.nascimento@lftm.com.br>; Willians Marques <willians.marques@lftm.com.br>'

if sendfromAC:
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, sendfromAC))
mail.To = 'Assessores (LIFETIME) <ASSESSORES@LIFETIMEINVEST.COM.BR>'
#mail.To = 'Tiago Sousa <tiago.sousa@LIFETIMEASSET.COM.BR>'
mail.Cc = cominterna
mail.Subject = 'Acompanhamento diário Balanced'
mail.Attachments.Add(r"Z:\GESTAO\FUNDOS ABERTOS\LIFETIME BALANCED\Lâminas\Lamina_Balanced para acompanhamento\Acompanhamento Diário Balanced.pdf")

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

if sendfromAC:
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, sendfromAC))
mail.To = 'Assessores (LIFETIME) <ASSESSORES@LIFETIMEINVEST.COM.BR>'
#mail.To = 'Tiago Sousa <tiago.sousa@LIFETIMEASSET.COM.BR>; andres.escobar@LIFETIMEASSET.COM.BR'
mail.Cc = 'Operations <operations@LIFETIMEASSET.COM.BR>; Comunicação Interna <comunicacaointerna@lftm.com.br>'
mail.Subject = 'Acompanhamento diário Graphene'
mail.Attachments.Add(r"Z:\GESTAO\FUNDOS ABERTOS\LIFETIME GRAPHENE\Lâmina\Lamina_Graphene para acompanhamento\Acompanhamento Diário Graphene.pdf")

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