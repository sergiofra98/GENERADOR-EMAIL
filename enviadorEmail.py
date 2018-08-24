# coding=utf-8
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders
import pdfkit
import os


# me == my email address
# you == recipient's email address
me = "sergio.franco@fygsolutions.com"
you = "sergio.franco@fygsolutions.com"

# Create message container - the correct MIME type is multipart/alternative.
msg = MIMEMultipart('alternative')
msg['Subject'] = "Link"
msg['From'] = me
msg['To'] = you

# Create the body of the message (a plain-text and an HTML version).
text = "Hi!\nHow are you?\nHere is the link you wanted:\nhttp://www.python.org"
html = """\
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style>
        body {
            font-family: Verdana, Geneva, Tahoma, sans-serif;
        }

        table td,
        table th {
            padding: .75rem;
            vertical-align: top;
            border-top: 1px solid #dee2e6;
            background-color: #fff;
        }

        thead th {
            vertical-align: bottom;
            color: #fff;
            background-color: #212529;
            border-color: #32383e
        }
    </style>
</head>

<body>
    <div class="header">
        <div class="col">
            <h4>17/08/2018</h4>

        </div>
        <div class="col">
            <b>Proyeccion:&nbsp;</b> 0.85
            <b>Al dia:&nbsp;</b> 22
            <b>Dias en el mes:&nbsp;</b> 22
        </div>
    </div>
    <div class="tabla">
        <h4>COLOCACIÓN POR DIVISIÓN</h4>
        <table class="table">
            <thead>
                <tr class="tituloTabla">
                    <th>División </th>
                    <th>Tubo </th>
                    <th>DISPUESTO</th>
                    <th>POR DISPONER</th>
                    <th>Total </th>
                    <th>Proyección </th>
                    <th>Objetivo </th>
                    <th>% </th>
                    <th>jul-17 </th>
                    <th>VS MAA </th>
                    <th>Acum/17 </th>
                    <th>Acum/18 </th>
                    <th>% Variación </th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
            </tbody>
        </table>
    </div>
    <div class="tabla">
        <h4>COLOCACIÓN POR DIVISIÓN</h4>
        <table class="table">
            <thead>
                <tr class="tituloTabla">
                    <th>División </th>
                    <th>Tubo </th>
                    <th>DISPUESTO</th>
                    <th>POR DISPONER</th>
                    <th>Total </th>
                    <th>Proyección </th>
                    <th>Objetivo </th>
                    <th>% </th>
                    <th>jul-17 </th>
                    <th>VS MAA </th>
                    <th>Acum/17 </th>
                    <th>Acum/18 </th>
                    <th>% Variación </th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
            </tbody>
        </table>
    </div>
    <div class="tabla">
        <h4>COLOCACIÓN POR DIVISIÓN</h4>
        <table class="table">
            <thead>
                <tr class="tituloTabla">
                    <th>División </th>
                    <th>Tubo </th>
                    <th>DISPUESTO</th>
                    <th>POR DISPONER</th>
                    <th>Total </th>
                    <th>Proyección </th>
                    <th>Objetivo </th>
                    <th>% </th>
                    <th>jul-17 </th>
                    <th>VS MAA </th>
                    <th>Acum/17 </th>
                    <th>Acum/18 </th>
                    <th>% Variación </th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
                <tr>
                    <td>VM SUR (JC)</td>
                    <td>290,909 </td>
                    <td>4,460,106 </td>
                    <td>127,786 </td>
                    <td>4,878,801 </td>
                    <td>4,460,106 </td>
                    <td>7,900,000 </td>
                    <td>-44% </td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                </tr>
            </tbody>
        </table>
    </div>
</body>

</html>
"""

# Record the MIME types of both parts - text/plain and text/html.
part1 = MIMEText(text, 'plain')
part2 = MIMEText(html, 'html')

pdfkit.from_string(html.decode('utf-8'), 'temp/Colocacion.pdf')

filename = "Colocacion.pdf"
attachment = open("temp/Colocacion.pdf", "rb")
part = MIMEBase('application', 'octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
# Attach parts into message container.
# According to RFC 2046, the last part of a multipart message, in this case
# the HTML message, is best and preferred.
msg.attach(part1)
msg.attach(part2)
msg.attach(part)
# Send the message via local SMTP server.
mail = smtplib.SMTP('smtp.gmail.com', 587)

mail.ehlo()

mail.starttls()

mail.login("sergio.franco@fygsolutions.com", "xxxxxxxx")

mail.sendmail(me, you, msg.as_string())
mail.quit()

attachment.close()

if os.path.exists("temp/Colocacion.pdf"):
  os.remove("temp/Colocacion.pdf")
else:
  print("The file does not exist")
