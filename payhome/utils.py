#! /usr/bin/env python
# encoding:utf-8
"""
Copyright (C) 2018 Yunrong Technology

description:
author:yutingting
time:2018/4/25
PN:
"""
import ConfigParser
def get_config(service_conf, section=''):
    config = ConfigParser.ConfigParser()
    config.read(service_conf)

    conf_items = dict(config.items('common')) if config.has_section('common') else {}
    if section and config.has_section(section):
       conf_items.update(config.items(section))
    return conf_items