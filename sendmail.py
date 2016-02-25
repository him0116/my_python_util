# -*- coding: utf-8 -*-
from __future__ import print_function

import os
import datetime
import getpass
import base64

import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

RECEIVER = {
}


def __send_mail(account, passwd, options):
    if 'text_content' not in options or 'html_content' not in options:
        log.problem('No content for email')
        return

    try:
        # smtp_server = smtplib.SMTP('localhost')
        smtp_server = smtplib.SMTP('smtp.gmail.com:587')
        smtp_server.ehlo()
        smtp_server.starttls()
        smtp_server.login(account, passwd)
    except smtplib.SMTPException:
        log.exception()
        return
    except smtplib.socket.error:
        # Todo : ????
        log.exception()
        return

    subject = options.get(
        'subject', u'[行政] %s 薪資單' % options['date'].strftime("%Y-%m"))
    email_from = options.get('from', '')
    email_to = options.get('to', '')
    text_content = options.get('text_content', '')
    html_content = options.get('html_content', '')

    # msg = MIMEText(content, _charset='utf8')
    try:
        msg_body = MIMEMultipart('alternative')

        part1 = MIMEText(text_content.encode('utf-8'), 'plain', 'utf-8')
        part2 = MIMEText(html_content.encode('utf-8'), 'html', 'utf-8')

        msg_body.attach(part1)
        msg_body.attach(part2)

        msg = MIMEMultipart('mixed')
        msg['Subject'] = subject
        msg['From'] = email_from
        msg['To'] = email_to
        msg.attach(msg_body)

        pdf_file = options.get('pdf_file', '')
        if pdf_file:
            # msg.attach(MIMEText(file(pdf_file).read()))
            print(u"sending {0} to {1}".format(pdf_file, email_to))
            with open(pdf_file, "rb") as fil:
                pdf_attachment = MIMEApplication(fil.read(), _subtype="pdf")
                pdf_attachment.add_header(
                    'Content-Disposition', 'attachment',
                    # filename=('utf-8', pdf_file, 'salary.pdf'))
                    filename='=?utf-8?b?' + base64.b64encode(
                        pdf_file.encode('UTF-8')) + '?=')
                msg.attach(pdf_attachment)
                '''
                msg.attach(MIMEApplication(
                    fil.read(),
                    Content_Disposition='attachment; filename="%s"' % pdf_file,
                    Name=pdf_file
                ))
                '''

        smtp_server.sendmail(email_from, email_to, msg.as_string())
    except BaseException as e:
        print(e)
        raise

    smtp_server.quit()


def sendmail(account='him0116', do_test=False):
    passwd = getpass.getpass("password of gmail %s\n" % account)

    text_tmpl = u'{name}，'
    text_tmpl += u'\n    薪資單請見附件，若有問題，請和我聯絡。'
    text_tmpl += u'\n\n雲灣資訊有限公司  '

    html_tmpl = u"""\
    <html>
      <head></head>
      <body>
        <p>{name}，<br>
           &nbsp;&nbsp;&nbsp;&nbsp;薪資單請見附件，若有問題，請和我聯絡。
           <br><br>
           雲灣資訊有限公司
        </p>
      </body>
    </html>
    """

    now = datetime.datetime.now()
    for employee, email in RECEIVER.iteritems():
        name = employee[1:]
        pdf_file = u'雲灣資訊_薪資單_{0}_{1}.pdf'.format(
            employee, now.strftime("%Y%m"))

        if do_test:
            mail_to = 'him@cloudybay.com.tw'
        else:
            mail_to = email

        if os.path.isfile(pdf_file):
            __send_mail(account, passwd, {
                'date': now,
                'to': mail_to,
                'text_content': text_tmpl.format(name=name),
                'html_content': html_tmpl.format(name=name),
                'pdf_file': pdf_file
            })
        else:
            print("Cannot find pdf file : ", pdf_file.encode('utf-8'))


if __name__ == "__main__":
    sendmail(do_test=True)
