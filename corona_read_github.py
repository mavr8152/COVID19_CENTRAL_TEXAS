# Import package
from urllib.request import urlretrieve

#set up emailer:
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart

port = 465  # For SSL
smtp_server = "smtp.gmail.com"
sender_email = "xxxx@xxx.com"  # Enter your address
receiver_email = "yyyy@yyy.com"  # Enter receiver address
password = 'Zaqwsx$123'  # Enter your password


# Assign url of file: url
url = 'https://dshs.texas.gov/coronavirus/TexasCOVID19CaseCountData.xlsx'

# Save file locally
urlretrieve(url, 'F:\Python\CORONA_DATA\CaseCountData.xlsx')

# Reading an excel file using Python 
import xlrd 

loc = ("F:\Python\CORONA_DATA\CaseCountData.xlsx") 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0) 
header = sheet.cell_value(0, 0)
h = header.split()
#print(h [0], h [9], h[10], h[11], h[12])
COV = header.split()[0]
DAT = header.split()[9]
TIM = header.split()[11]
TZ = header.split()[12]



    #  Extract first column
for i in range(sheet.nrows): 
    county = sheet.cell_value(i, 0)
    # row data:
    rowpositive = sheet.cell_value(i, 1)
    rowfatal = sheet.cell_value(i, 2)

    #extract county name
    if county == 'Travis' or county == 'Williamson' :
        print('For county ', county, 'Positive ', rowpositive, 'Fatalities ', rowfatal)

        message = MIMEMultipart("alternative")
        message["Subject"] = "Texas " + str(COV) + " update " + str(DAT) + " " + str(TIM)
        message["From"] = sender_email
        message["To"] = receiver_email

        # write the HTML part
        html = """\
        <html>
         <body>
            <p>""" + str(COV) + """ update for county : """ + str(county) + """<br></p>
            <p>Positive: """ + str(rowpositive) + """ Fatalities: """ + str(rowfatal) + """<br></p>
            </p>
            <img src="cid:CoronaData">
         </body>
        </html>
        """
        part = MIMEText(html, "html")
        message.attach(part)

        # attaching png images (extracted with VBA macro and running on windows scheduler before the python script)
        if county == 'Travis' :
            fp = open('F:\Python\CORONA_DATA\Travis.png', 'rb')
            image = MIMEImage(fp.read())
            fp.close()

        if county == 'Williamson' :
            fp = open('F:\Python\CORONA_DATA\Williamson.png', 'rb')
            image = MIMEImage(fp.read())
            fp.close()
        
        # Specify the  ID according to the img src in the HTML part
        image.add_header('Content-ID', '<CoronaData>')
        message.attach(image)

        # send your email
        with smtplib.SMTP_SSL(smtp_server, port) as server:
            server.login(sender_email, password)
            server.sendmail(
                sender_email, receiver_email, message.as_string()
            )

