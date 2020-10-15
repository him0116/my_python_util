import os
import sys
import datetime
import base64

from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import argparse
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

sys.path.append('/Users/him/workspace/site-packages')
try:
    from cb import log
except ImportError:
    class log(object):
        @classmethod
        def problem(cls, msg):
            print(msg)

        @classmethod
        def exception(cls, msg):
            print(msg)

RECEIVER = {
    u'吳柏漢': 'roywu@cloudybay.com.tw',
    u'郭倢汝': 'jocelyn@cloudybay.com.tw',
    # # u'江啟安': 'viserys@cloudybay.com.tw',
    # # u'趙訢懿': 'cindychao@cloudybay.com.tw',
    u'謝依蓉': 'neko@cloudybay.com.tw',
    u'劉佳琪': 'lorie@cloudybay.com.tw',
    # # u'劉若珊': 'mavisliu@cloudybay.com.tw',
    u'李哲維': 'stanleeley@cloudybay.com.tw',
    # # u'蘇胤瑞': 'yinruei@cloudybay.com.tw',
    u'陳品妤': 'pinyu@cloudybay.com.tw',
    u'康申': 'skang@cloudybay.com.tw',
    u'周彥伶': 'irenechou@cloudybay.com.tw',
    u'張少翔': 'ina@cloudybay.com.tw',
    u'李佳穎': 'cylee29@cloudybay.com.tw',
}

USE_LOCALHOST = False
SCOPES = 'https://www.googleapis.com/auth/gmail.send'


def create_message(options):
    if 'text_content' not in options or 'html_content' not in options:
        log.problem('No content for email')
        return None

    subject = options.get(
        'subject', u'[行政] %s 薪資單' % options['date'].strftime("%Y-%m"))
    email_from = options.get('from', 'him@cloudybay.com.tw')
    email_to = options.get('to', 'him@cloudybay.com.tw')
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
            with open(pdf_file, "rb") as fil:
                pdf_attachment = MIMEApplication(fil.read(), _subtype="pdf")
                pdf_attachment.add_header(
                    'Content-Disposition', 'attachment',
                    # filename=('utf-8', pdf_file, 'salary.pdf'))
                    filename=pdf_file)
                msg.attach(pdf_attachment)
                '''
                msg.attach(MIMEApplication(
                    fil.read(),
                    Content_Disposition='attachment; filename="%s"' % pdf_file,
                    Name=pdf_file
                ))
                '''
        return {'raw': base64.urlsafe_b64encode(msg.as_string().encode()).decode()}
        # return msg.as_string()
    except BaseException as e:
        print(e)
        raise


def get_gmail_service():
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('gmail', 'v1', http=creds.authorize(Http()))
    return service


def send_message(account, service, message):
    """Send an email message.

      Args:
        account: User's email address. The special value "me"
        can be used to indicate the authenticated user.
        service: Authorized Gmail API service instance.
        message: Message to be sent.

      Returns:
        Sent Message.
    """

    try:
        message = (service.users().messages().send(userId=account, body=message)
                   .execute())
        print('Message Id: %s' % message['id'])
        return message
    except BaseException as error:
        print('An error occurred: %s' % error)
        return


def sendmail(account='him0116@gmail.com', do_test=False, salary_date=None):
    # passwd = getpass.getpass("password of gmail %s\n" % account)

    text_tmpl = u'{name}，'
    text_tmpl += u'\n    薪資單請見附件，若有問題，請和我聯絡。'
    text_tmpl += u'\n\n雲灣資訊有限公司  洪一民'

    html_tmpl = u"""\
    <html>
      <head></head>
      <body>
        <p>{name}，<br>
           &nbsp;&nbsp;&nbsp;&nbsp;薪資單請見附件，若有問題，請和我聯絡。
           <br><br>
           雲灣資訊有限公司  洪一民
        </p>
      </body>
    </html>
    """

    service = get_gmail_service()

    if salary_date is None:
        now = datetime.datetime.now()
    else:
        now = salary_date

    for employee, email in RECEIVER.items():
        name = employee[1:] if len(employee) > 2 else employee
        pdf_file = u'雲灣資訊_薪資單_{0}_{1}.pdf'.format(
            employee, now.strftime("%Y%m"))

        if do_test:
            mail_to = 'him@cloudybay.com.tw'
        else:
            mail_to = email

        if os.path.isfile(pdf_file):
            options = {
                'date': now,
                'to': mail_to,
                'text_content': text_tmpl.format(name=name),
                'html_content': html_tmpl.format(name=name),
                'pdf_file': pdf_file
            }
            msg = create_message(options)
            if msg is not None:
                print(u"sending {0} to {1}".format(pdf_file, mail_to))
                send_message(account, service, msg)
        else:
            print("Cannot find pdf file : ", pdf_file.encode('utf-8'))


usage = """
  -s : execute send mail (do_test=False)
  -d : salary date (yyyymmdd)
  -h : show this message
"""

if __name__ == "__main__":
    # import getopt
    # opts, args = getopt.getopt(sys.argv[1:], "shd:")
    tools.argparser.add_argument('-s', '--send', required=False,
                                 action="store_true",
                                 help='execute send mail (do_test=False)')
    tools.argparser.add_argument('-d', '--date', type=str, required=False,
                                 help='salary date')
    # option = ""
    '''
    for opt, arg in opts:
        if opt == "-s":
            do_test = False
        if opt == "-d":
            try:
                salary_date = datetime.datetime.strptime(arg, "%Y%m%d")
            except e:
                print(e)
        if opt == "-h":
            print(usage)
            sys.exit(0)
    '''
    args = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
    print(args)
    do_test = not args.send
    if args.date is None:
        salary_date = datetime.datetime.now()
    else:
        salary_date = datetime.datetime.strptime(args.date, "%Y%m%d")

    sendmail(do_test=do_test, salary_date=salary_date)
