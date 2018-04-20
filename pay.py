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
    data = xlrd.open_workbook(file_path)
    table = data.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols
    data = []
    for i in xrange(0, nrows):
        rowValues = table.row_values(i)
        for item in rowValues:
            data = item

    send_email(data)
    os.remove(file_path)

def send_email(data):

    # 输入Email地址和口令:
    from_addr = 'yutingting@yunrongtech.com'
    password = 'yutingting321'
    # 输入SMTP服务器地址:
    smtp_server = 'mail.yunrongtech.com'
    # 输入收件人地址:
    to_addr = 'yutingting@yunrongtech.com'

    msg = MIMEText('hello, send by my heart...', 'plain', 'utf-8')
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
