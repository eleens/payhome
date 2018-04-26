# encoding: utf-8
import os
from flask import render_template,redirect, url_for
from flask_uploads import UploadSet, configure_uploads, DOCUMENTS, patch_request_class
import xlrd
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr
import smtplib
from email.header import Header
from . import app
from payhome.web.forms import *


work_dir = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
upload_path = os.path.join(work_dir, 'uploads')
app.config['UPLOADED_PHOTOS_DEST'] = upload_path
files = UploadSet('photos', DOCUMENTS)
configure_uploads(app, files)
patch_request_class(app)

@app.route('/', methods=['GET', 'POST'])
def index():
    form = FileForm()
    if form.validate_on_submit():
        filename = files.save(form.files.data)
        file_path = files.path(filename)
        handle_excel(file_path)
        return redirect(url_for('index'))
    return render_template('index.html', form=form)


def handle_excel(file_path):
    excel = xlrd.open_workbook(file_path)
    table = excel.sheets()[0]
    nrows = table.nrows
    name = table.name
    fields = {}
    title = name
    email_data = {
        'user_name': unicode(app.config.get('SMTP_USER_NAME'),'utf-8'),
        'user_dep': unicode(app.config.get('SMTP_USER_DEP'), 'utf-8'),
        'user_email': app.config.get('SMTP_USER_EMAIL'),
        'company_ename': app.config.get('COMPANY_ENAME'),
        'company_name': unicode(app.config.get('COMPANY_NAME'), 'utf-8'),
        'company_tel': app.config.get('COMPANY_TEL'),
        'company_fax': app.config.get('COMPANY_FAX'),
        'company_addr': unicode(app.config.get('COMPANY_ADDR'), 'utf-8'),
        'user_tel':  app.config.get('SMTP_USER_TEL'),
        'company_post': app.config.get('COMPANY_POST')
    }


    for i in xrange(0, nrows):
        rowValues = table.row_values(i)
        pays = {}
        for index, item in enumerate(rowValues):
            if i == 0:
                title = item or title
            elif i == 1:
                fields[index] = item
            else:
                pays[fields.get(index)] = item
        if i not in [0, 1]:
            e_mail = pays.pop('email')
            tel = pays.pop('tel')
            field_list = fields.values()
            field_list.remove('email')
            field_list.remove('tel')
            result = render_template('email.html',
                                     email_data=email_data,
                                     pay_title=title,
                                     field_list=field_list,
                                     datas=pays)
            tel_result = ''
            if e_mail and result:
                send_email(name, result, e_mail)
                print "send"
            if tel and tel_result:
                send_message(tel_result, tel)

    os.remove(file_path)

def send_message(tel_result, tel):
    pass

def send_email(theme, result, e_mail):

    # 输入Email地址和口令:
    from_addr = app.config.get('SMTP_USER_EMAIL')
    password = app.config.get('SMTP_USER_PASS')
    # 输入SMTP服务器地址:
    smtp_server = app.config.get('SMTP_SERVER')
    # 输入收件人地址:
    to_addr = e_mail


    msg = MIMEText(result, 'html', 'utf-8')
    msg['From'] = _format_addr('%s <%s>' % (from_addr, from_addr))
    msg['To'] = _format_addr('%s <%s>' % (to_addr, to_addr))
    msg['Subject'] = Header(theme, 'utf-8').encode()


    server = smtplib.SMTP(smtp_server, 25)  # SMTP协议默认端口是25
    server.set_debuglevel(1)
    server.login(from_addr, password)
    server.sendmail(from_addr, [to_addr], msg.as_string())
    server.quit()

def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((
        Header(name, 'utf-8').encode(),
        addr.encode('utf-8') if isinstance(addr, unicode) else addr))


if __name__ == '__main__':
    app.run()
