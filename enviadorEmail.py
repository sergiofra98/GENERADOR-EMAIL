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
you = "mario.mejorada@fygsolutions.com"

# Create message container - the correct MIME type is multipart/alternative.
msg = MIMEMultipart('alternative')
msg['Subject'] = "Resumen Diario Por Division"
msg['From'] = me
msg['To'] = you

# Create the body of the message (a plain-text and an HTML version).
text = "Resumen Diario Por Division"
html = """\
<!DOCTYPE html>
<html>

<head>
    <style>
        body {
            font-family: Verdana, Geneva, Tahoma, sans-serif;
        }

        table {
            border: 1px solid #dee2e6;
            width: 100%;
        }

        .header div {
            width: 50%;
            float: left;
        }

        .header div:last-of-type{
            text-align: right;
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
            background-color: #0080ff;
            border-color: #0056ac;
            border-bottom-width: 2px
        }

        .tituloObscuro th {
            background-color: #445261;
        }

        td {
            text-align: right;
            border-left: 1px solid #dee2e6;
        }

        #tablaDivision tr td:nth-of-type(4),
        #tablaDivision tr td:nth-of-type(10),
        #tablaDivision tr td:nth-of-type(13),
        #tablaDivision tr td:nth-of-type(14) {
            background-color: #abd5ff;
        }

        #tablaDivision tr td:nth-of-type(5) {
            background-color: #0080ff;
            color: #fff;
        }

        td:first-of-type {
            text-align: left;
            border-left: none;
        }

        .rojo {
            background-color: red;
        }
    </style>
    <script>
    </script>
</head>

<body>
    <div class="header">
        <div>
            <b>Al 31 de julio de 2018</b>
        </div>
        <div>
            <b>Proyección:&nbsp;</b>0.85<br>
            <b>Al día:&nbsp;</b>22 <br>
            <b>Días en el mes:&nbsp;</b>22 <br>
        </div>
    </div>
    <div class="tabla">
        <h4>COLOCACIÓN POR DIVISIÓN</h4>
        <table id="tablaDivision">
            <thead>
                <tr class="tituloObscuro">
                    <th colspan="9">Colocación Mes Actual</th>
                    <th colspan="3">Colocación VS</th>
                    <th colspan="4">Colocación 2018 VS AA</th>
                </tr>
                <tr class="tituloTabla">
                    <th>División </th>
                    <th>TUBO </th>
                    <th>DISPUESTO</th>
                    <th>POR DISPONER</th>
                    <th>TOTAL </th>
                    <th>PROYECCIÓN </th>
                    <th>OBJETIVO </th>
                    <th colspan="2">% </th>
                    <th>JUL 17 </th>
                    <th colspan="2">VS MAA </th>
                    <th>ACUM/17 </th>
                    <th>ACUM/18 </th>
                    <th colspan="2">% VARIACIÓN </th>
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
                    <td class="rojo"></td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td class="rojo"></td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                    <td class="rojo"></td>

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
                    <td class="rojo"></td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td class="rojo"></td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                    <td class="rojo"></td>

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
                    <td class="rojo"></td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td class="rojo"></td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                    <td class="rojo"></td>

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
                    <td class="rojo"></td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td class="rojo"></td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                    <td class="rojo"></td>

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
                    <td class="rojo"></td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td class="rojo"></td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                    <td class="rojo"></td>

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
                    <td class="rojo"></td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td class="rojo"></td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                    <td class="rojo"></td>

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
                    <td class="rojo"></td>
                    <td>5,276,908 </td>
                    <td>-15% </td>
                    <td class="rojo"></td>
                    <td>33,375,027 </td>
                    <td>29,090,407 </td>
                    <td>-13% </td>
                    <td class="rojo"></td>

                </tr>
            </tbody>
        </table>
    </div>
    <div class="tabla">
        <h4>COLOCACIÓN POR CONVENIO</h4>
        <table id="tablaConvenio">
            <thead>
                <tr class="tituloObscuro">
                    <th colspan="5">VS Mes Año Anterior</th>
                    <th colspan="4">Acumulado Año</th>
                    <th colspan="4">Colocación Por Convenio</th>
                </tr>
                <tr class="tituloTabla">
                    <th>División </th>
                    <th>JUN 17 </th>
                    <th>JUN 18</th>
                    <th colspan="2">VS MAA</th>
                    <th>JUN 17 </th>
                    <th>JUN 18</th>
                    <th colspan="2">% Var</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>GDF</td>
                    <td>13,313,692</td>
                    <td>10,306,536</td>
                    <td>-23%</td>
                    <td class="rojo"></td>
                    <td>74,536,396</td>
                    <td>68,453,360</td>
                    <td>-8%</td>
                    <td class="rojo"></td>
                </tr><tr>
                    <td>GDF</td>
                    <td>13,313,692</td>
                    <td>10,306,536</td>
                    <td>-23%</td>
                    <td class="rojo"></td>
                    <td>74,536,396</td>
                    <td>68,453,360</td>
                    <td>-8%</td>
                    <td class="rojo"></td>
                </tr><tr>
                    <td>GDF</td>
                    <td>13,313,692</td>
                    <td>10,306,536</td>
                    <td>-23%</td>
                    <td class="rojo"></td>
                    <td>74,536,396</td>
                    <td>68,453,360</td>
                    <td>-8%</td>
                    <td class="rojo"></td>
                </tr><tr>
                    <td>GDF</td>
                    <td>13,313,692</td>
                    <td>10,306,536</td>
                    <td>-23%</td>
                    <td class="rojo"></td>
                    <td>74,536,396</td>
                    <td>68,453,360</td>
                    <td>-8%</td>
                    <td class="rojo"></td>
                </tr><tr>
                    <td>GDF</td>
                    <td>13,313,692</td>
                    <td>10,306,536</td>
                    <td>-23%</td>
                    <td class="rojo"></td>
                    <td>74,536,396</td>
                    <td>68,453,360</td>
                    <td>-8%</td>
                    <td class="rojo"></td>
                </tr><tr>
                    <td>GDF</td>
                    <td>13,313,692</td>
                    <td>10,306,536</td>
                    <td>-23%</td>
                    <td class="rojo"></td>
                    <td>74,536,396</td>
                    <td>68,453,360</td>
                    <td>-8%</td>
                    <td class="rojo"></td>
                </tr><tr>
                    <td>GDF</td>
                    <td>13,313,692</td>
                    <td>10,306,536</td>
                    <td>-23%</td>
                    <td class="rojo"></td>
                    <td>74,536,396</td>
                    <td>68,453,360</td>
                    <td>-8%</td>
                    <td class="rojo"></td>
                </tr>
            </tbody>
        </table>
    </div>
    <div class="tabla">
        <h4>PLANTILLA</h4>
        <table id="tablaPlantilla">
            <thead>
                <tr class="tituloObscuro">
                    <th colspan="5">Productividad por asesor</th>
                    <th colspan="2">Plantilla</th>
                    <th colspan="2">Rotación</th>
                    <th colspan="4">Asesores Improductivos</th>
                </tr>
                <tr class="tituloTabla">
                    <th>División </th>
                    <th>Plantilla</th>
                    <th>Venta Promedio por Asesor</th>
                    <th>Improductivos >2M</th>
                    <th>% de Improductivos</th>
                    <th>HC Autorizado</th>
                    <th>Cubrimiento</th>
                    <th>Rotación Acum 2017</th>
                    <th>Rotación Acum 2018</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>VM SUR</td>
                    <td>64</td>
                    <td>69,689</td>
                    <td>2</td>
                    <td>3%</td>
                    <td>72 </td>
                    <td>89%</td>
                    <td>52%</td>
                    <td>59%</td>

                </tr><tr>
                    <td>VM SUR</td>
                    <td>64</td>
                    <td>69,689</td>
                    <td>2</td>
                    <td>3%</td>
                    <td>72 </td>
                    <td>89%</td>
                    <td>52%</td>
                    <td>59%</td>

                </tr><tr>
                    <td>VM SUR</td>
                    <td>64</td>
                    <td>69,689</td>
                    <td>2</td>
                    <td>3%</td>
                    <td>72 </td>
                    <td>89%</td>
                    <td>52%</td>
                    <td>59%</td>

                </tr><tr>
                    <td>VM SUR</td>
                    <td>64</td>
                    <td>69,689</td>
                    <td>2</td>
                    <td>3%</td>
                    <td>72 </td>
                    <td>89%</td>
                    <td>52%</td>
                    <td>59%</td>

                </tr><tr>
                    <td>VM SUR</td>
                    <td>64</td>
                    <td>69,689</td>
                    <td>2</td>
                    <td>3%</td>
                    <td>72 </td>
                    <td>89%</td>
                    <td>52%</td>
                    <td>59%</td>

                </tr><tr>
                    <td>VM SUR</td>
                    <td>64</td>
                    <td>69,689</td>
                    <td>2</td>
                    <td>3%</td>
                    <td>72 </td>
                    <td>89%</td>
                    <td>52%</td>
                    <td>59%</td>

                </tr><tr>
                    <td>VM SUR</td>
                    <td>64</td>
                    <td>69,689</td>
                    <td>2</td>
                    <td>3%</td>
                    <td>72 </td>
                    <td>89%</td>
                    <td>52%</td>
                    <td>59%</td>

                </tr><tr>
                    <td>VM SUR</td>
                    <td>64</td>
                    <td>69,689</td>
                    <td>2</td>
                    <td>3%</td>
                    <td>72 </td>
                    <td>89%</td>
                    <td>52%</td>
                    <td>59%</td>

                </tr>
            </tbody>
        </table>
    </div>
    <div class="tabla">
        <h4>SUPERVISORES</h4>
        <table id="tablaSupervisor">
            <thead>
                <tr class="tituloTabla">
                    <th>Supervisor</th>
                    <th>No. Persona</th>
                    <th>Asesor</th>
                    <th>Sucursal</th>
                    <th>Fecha Ingreso</th>
                    <th>Jul 18</th>
                    <th>Jul 18</th>
                    <th>Ene 18</th>
                    <th>Feb 18</th>
                    <th>Mar 18</th>
                    <th>Abr 18</th>
                    <th>May 18</th>
                    <th>Jun 18</th>
                    <th colspan="2">Promedio</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>LAURA DENISSE RODRIGUEZ ALVISO</td>
                    <td>70525801</td>
                    <td>MARIA EUGENIA MALAGON CHAGOYA</td>
                    <td>GUANAJUATO</td>
                    <td>12/6/17</td>
                    <td class="rojo"></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td class="rojo"></td>

                </tr><tr>
                    <td>LAURA DENISSE RODRIGUEZ ALVISO</td>
                    <td>70525801</td>
                    <td>MARIA EUGENIA MALAGON CHAGOYA</td>
                    <td>GUANAJUATO</td>
                    <td>12/6/17</td>
                    <td class="rojo"></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td class="rojo"></td>

                </tr><tr>
                    <td>LAURA DENISSE RODRIGUEZ ALVISO</td>
                    <td>70525801</td>
                    <td>MARIA EUGENIA MALAGON CHAGOYA</td>
                    <td>GUANAJUATO</td>
                    <td>12/6/17</td>
                    <td class="rojo"></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td class="rojo"></td>

                </tr><tr>
                    <td>LAURA DENISSE RODRIGUEZ ALVISO</td>
                    <td>70525801</td>
                    <td>MARIA EUGENIA MALAGON CHAGOYA</td>
                    <td>GUANAJUATO</td>
                    <td>12/6/17</td>
                    <td class="rojo"></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td class="rojo"></td>

                </tr><tr>
                    <td>LAURA DENISSE RODRIGUEZ ALVISO</td>
                    <td>70525801</td>
                    <td>MARIA EUGENIA MALAGON CHAGOYA</td>
                    <td>GUANAJUATO</td>
                    <td>12/6/17</td>
                    <td class="rojo"></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td class="rojo"></td>

                </tr><tr>
                    <td>LAURA DENISSE RODRIGUEZ ALVISO</td>
                    <td>70525801</td>
                    <td>MARIA EUGENIA MALAGON CHAGOYA</td>
                    <td>GUANAJUATO</td>
                    <td>12/6/17</td>
                    <td class="rojo"></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td class="rojo"></td>

                </tr><tr>
                    <td>LAURA DENISSE RODRIGUEZ ALVISO</td>
                    <td>70525801</td>
                    <td>MARIA EUGENIA MALAGON CHAGOYA</td>
                    <td>GUANAJUATO</td>
                    <td>12/6/17</td>
                    <td class="rojo"></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td class="rojo"></td>

                </tr><tr>
                    <td>LAURA DENISSE RODRIGUEZ ALVISO</td>
                    <td>70525801</td>
                    <td>MARIA EUGENIA MALAGON CHAGOYA</td>
                    <td>GUANAJUATO</td>
                    <td>12/6/17</td>
                    <td class="rojo"></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td class="rojo"></td>

                </tr><tr>
                    <td>LAURA DENISSE RODRIGUEZ ALVISO</td>
                    <td>70525801</td>
                    <td>MARIA EUGENIA MALAGON CHAGOYA</td>
                    <td>GUANAJUATO</td>
                    <td>12/6/17</td>
                    <td class="rojo"></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td class="rojo"></td>

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

mail.login("sergio.franco@fygsolutions.com", "")

mail.sendmail(me, you, msg.as_string())
mail.quit()

attachment.close()

if os.path.exists("temp/Colocacion.pdf"):
  os.remove("temp/Colocacion.pdf")
else:
  print("The file does not exist")
