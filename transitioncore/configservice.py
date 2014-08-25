# ------------------------------------------------------------------------------
# Name:        configservice.py
# Purpose:     Expose config API via RPyC server
#
# Author:      Jonathan Besanceney <jonathan.besanceney@gmail.com>
#
# Created:     01/06/14
# Copyright:   (c) 2014 Jonathan Besanceney
#
# This file is a part of Transition
#
#    Transition is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    Transition is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU Lesser General Public License for more details.
#
#    You should have received a copy of the GNU Lesser General Public License
#    along with Transition.  If not, see <http://www.gnu.org/licenses/>.
# ------------------------------------------------------------------------------
# -*- coding: utf8 -*-

import inspect

import pickle
import os
from os.path import expanduser
import pkgutil
import sys

import rpyc

from transitioncore import TransitionAppType, transition_app_path

sys.path.append(os.path.abspath(os.path.dirname(__file__)))

from transitioncore.configuration import Configuration


class ConfigService(rpyc.Service):
    """
    Expose Configuration API with RPyC
    """

    def on_connect(self):
        self._config = Configuration()

    def exposed_get_app_list(self, app_type=None):
        return self._config.get_app_list(app_type)

    def exposed_disable_app(self, app_type, app_name):
        return self._config.disable_app(app_type, app_name)

    def exposed_enable_app(self, app_type, app_name):
        return self._config.enable_app(app_type, app_name)

    def exposed_get_app_desc(self, app_type, app_name):
        return self._config.get_app_desc(app_type, app_name)

    def exposed_get_app_status(self, app_type, app_name):
        return self._config.get_app_status(app_type, app_name)

    def exposed_get_disabled_app_list(self, app_type):
        return self._config.get_disabled_app_list(app_type)

    def exposed_get_enabled_app_list(self, app_type):
        return self._config.get_enabled_app_list(app_type)

    def exposed_add_event_listener(self, listener):
        print(repr(listener))
        self._config.add_event_listener(listener)

    def exposed_del_event_listener(self, listener):
        self._config.del_event_listener(listener)

if __name__ == "__main__":
    from rpyc.utils.server import ThreadedServer
    ThreadedServer(ConfigService, port=22).start()
