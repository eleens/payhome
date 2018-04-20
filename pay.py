# encoding: utf-8
import os
from flask import Flask
from flask import render_template,redirect, url_for
from flask_bootstrap import Bootstrap
from flask_wtf import FlaskForm
from flask_wtf.file import FileField, FileRequired
from wtforms import SubmitField
from flask_uploads import UploadSet, configure_uploads, DOCUMENTS, patch_request_class
import xlrd
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr
import smtplib
from email.header import Header

app = Flask(__name__)
bootstrap = Bootstrap(app)
app.config['SECRET_KEY']='xxx'



app.config['UPLOADED_PHOTOS_DEST'] = os.path.join(os.getcwd(), 'uploads')
photos = UploadSet('photos', DOCUMENTS)
configure_uploads(app, photos)
patch_request_class(app)

class PhotoForm(FlaskForm):
    photo = FileField(validators=[FileRequired()])
    submit = SubmitField('submit')


@app.route('/', methods=['GET', 'POST'])
def index():
    form = PhotoForm()
    if form.validate_on_submit():
        filename = photos.save(form.photo.data)
        file_path = photos.path(filename)
        handle_excel(file_path)
        return redirect(url_for('index'))
    return render_template('index.html', form=form)


def handle_excel(file_path):
    excel = xlrd.open_workbook(file_path)
    table = excel.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols
    fields = {}
    datas = []
    title = ''
    result = ''

    for i in xrange(0, nrows):
        rowValues = table.row_values(i)
        pays = {}
        for item in rowValues:
            if i == 0:
                title = item
            elif i in [1]:
                if item:
                    fields[rowValues.index(item)] = item
            else:
                pays[fields.get(rowValues.index(item))] = item
        if i not in [0, 1]:
            other_pays = list(set(fields.values()) - set(pays.keys()))
            e_mail = pays.pop('email')
            tel = pays.pop('tel')
            for x in other_pays:
                pays[x] = '-'
            for f, p in pays.items():
                re = "<tr><th>%s</th><td>%s</td></tr>" % (f, p)
                datas.append(re)
            result = render_template('email.html', title=title, datas=datas)
            send_email(result, e_mail)
            break

    os.remove(file_path)

def send_email(result, e_mail):

    # 输入Email地址和口令:
    from_addr = 'yutingting@yunrongtech.com'
    password = 'yutingting321'
    # 输入SMTP服务器地址:
    smtp_server = 'mail.yunrongtech.com'
    # 输入收件人地址:
    # to_addr = 'yutingting@yunrongtech.com'
    to_addr = e_mail

    msg = MIMEText(result, 'html', 'utf-8')
    msg['From'] = _format_addr('Love You <%s>' % from_addr)
    msg['To'] = _format_addr('My Love <%s>' % to_addr)
    msg['Subject'] = Header('Love', 'utf-8').encode()


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
