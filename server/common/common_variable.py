# -*- coding: utf8 -*-

import os
from common import common_definition as CMN_DEF


class GlobalVar(object):

    _GLOBAL_VARIABLE_UPDATED = False
    _HOSTNAME = "localhost"

    class __metaclass__(type):
        @property
        def GLOBAL_VARIABLE_UPDATED(cls):
            raise RuntimeError("Can NOT be invoked")


        @GLOBAL_VARIABLE_UPDATED.setter
        def GLOBAL_VARIABLE_UPDATED(cls, global_variable_updated):
            if cls._GLOBAL_VARIABLE_UPDATED:
                raise RuntimeError("Global variables have already been UPDATED !!!")
            cls._GLOBAL_VARIABLE_UPDATED = global_variable_updated


        @property
        def HOSTNAME(cls):
            if not cls._GLOBAL_VARIABLE_UPDATED:
                raise RuntimeError("Global variables are NOT updated !!!")
            return cls._HOSTNAME


        @HOSTNAME.setter
        def HOSTNAME(cls, hostname):
            if cls._GLOBAL_VARIABLE_UPDATED:
                raise RuntimeError("Global variables have already been UPDATED !!!")
            cls._HOSTNAME = hostname

