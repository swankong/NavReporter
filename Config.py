# -*- coding: utf-8 -*-

#
# Config.py
#
# Load and parse yaml config
#
# All rights reserved, 2019
#

import yaml
import collections
import six
import pprint

class AttrDict(object):
    """ 将配置属性字典映射为成员结构体形式 """

    def __init__(self, d=None):
        self.__dict__ = d if d is not None else dict()

        for k, v in list(six.iteritems(self.__dict__)):
            if isinstance(v, dict):
                self.__dict__[k] = AttrDict(v)  # 拆解成成员，放到self.__dict__结构中

    def __repr__(self):
        return pprint.pformat(self.__dict__)

    def __iter__(self):
        return self.__dict__.__iter__()

    def update(self, other):
        AttrDict.__update_dict_recursive(self, other)

    def items(self):
        return six.iteritems(self.__dict__)

    iteritems = items

    def keys(self):
        return self.__dict__.keys()

    @staticmethod
    def __update_dict_recursive(target, other):
        if isinstance(other, AttrDict):
            other = other.__dict__
        if isinstance(target, AttrDict):
            target = target.__dict__

        for k, v in six.iteritems(other):
            if isinstance(v, collections.Mapping):
                r = AttrDict.__update_dict_recursive(target.get(k, {}), v)
                target[k] = r
            else:
                target[k] = other[k]
        return target

    def convert_to_dict(self):
        result_dict = {}
        for k, v in list(six.iteritems(self.__dict__)):
            if isinstance(v, AttrDict):
                v = v.convert_to_dict()
            result_dict[k] = v
        return result_dict

class Config(object):

    def __init__(self, config_file=None):
        self.conf_file = config_file
        self.conf_dict = None

    def load_config(self):
        with open(self.conf_file, 'r', encoding='UTF-8') as f:
            self.conf_dict = yaml.load(f.read(), Loader=yaml.BaseLoader)

    def parse_config(self):
        conf_ = AttrDict(self.conf_dict)
        return conf_

    def get_config(self, conf_name):
        return self.conf_dict.get(conf_name, None)



