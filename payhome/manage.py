#! /usr/bin/env python
# encoding: utf-8
"""
Copyright (C) 2018 Yunrong Technology

description：
author：yutingting
time：2018/4/25
PN: 
"""
import os
from web import app
from flask_bootstrap import Bootstrap
from flask_uploads import UploadSet, configure_uploads, DOCUMENTS, patch_request_class

work_dir = os.path.dirname(os.path.realpath(__file__))
config_path = os.path.join(work_dir, 'config.py')

bootstrap = Bootstrap(app)
app.config.from_pyfile(config_path)


host = app.config.get('HOST', '0.0.0.0')
port = app.config.get('PORT')

if __name__ == '__main__':
    app.run(host=host, port=port)