#! /usr/bin/env python
# encoding: utf-8
"""
Copyright (C) 2018 Yunrong Technology

description：
author：yutingting
time：2018/4/25
PN: 
"""
from flask_wtf import FlaskForm
from flask_wtf.file import FileField, FileRequired
from wtforms import SubmitField


class FileForm(FlaskForm):
    files = FileField(validators=[FileRequired()])
    submit = SubmitField('submit')
