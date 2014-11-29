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

import os
import sys

import rpyc

sys.path.append(os.path.abspath(os.path.dirname(__file__)))

from transitioncore.configuration import Configuration
from transitioncore.eventdispatcher import TransitionEventDispatcher
from transitioncore.eventsinterface.configeventinterface import ConfigEventsInterface

class ConfigService(rpyc.Service, ConfigEventsInterface, TransitionEventDispatcher):
    """
    Expose Configuration API with RPyC.
    This represents one config service Thread.
    """

    _config = None

    @staticmethod
    def config(config):
        ConfigService._config = config

    def on_connect(self):
        super(TransitionEventDispatcher, self).__init__()
        ConfigService._config.add_event_listener(self)

    def on_disconnect(self):
        ConfigService._config.del_event_listener(self)
        ConfigService._config = None

    def on_app_add(self, app_type, app_name):
        print('on_app_add')
        self._fire_event('on_app_add', (app_type, app_name))

    def on_app_del(self, app_type, app_name):
        print('on_app_del')

    def on_app_disable(self, app_type, app_name, com_app_tuple):
        print('on_app_disable')

    def on_app_update(self, app_type, app_name):
        print('on_app_update')

    def on_app_enable(self, app_type, app_name, com_app_tuple):
        print('on_app_enable')

    def add_event_listener(self, listener):
        pass

    def del_event_listener(self, listener):
        pass

    @staticmethod
    def exposed_get_app_list(app_type=None):
        return ConfigService._config.get_app_list(app_type)

    @staticmethod
    def exposed_disable_app(app_type, app_name):
        return ConfigService._config.disable_app(app_type, app_name)

    @staticmethod
    def exposed_enable_app(app_type, app_name):
        return ConfigService._config.enable_app(app_type, app_name)

    @staticmethod
    def exposed_get_app_desc(app_type, app_name):
        return ConfigService._config._get_app_desc(app_type, app_name)

    @staticmethod
    def exposed_get_disabled_app_list(app_type, com_app):
        return ConfigService._config.get_disabled_app_list(app_type, com_app)

    @staticmethod
    def exposed_get_enabled_app_list(app_type, com_app):
        return ConfigService._config.get_enabled_app_list(app_type, com_app)

    def exposed_add_event_listener(self, listener):
        self.listener_list.append(listener)
        ConfigService._config.add_event_listener(listener)

    def exposed_del_event_listener(self, listener):
        ConfigService._config.del_event_listener(listener)
        self.listener_list.remove(listener)

if __name__ == "__main__":
    from rpyc.utils.server import ThreadedServer
    ts = ThreadedServer(ConfigService, port=22, protocol_config={'allow_public_attrs': True,
                                                            'allow_all_attrs': True,
                                                            'allow_pickle': True,
                                                            'allow_setattr': True,
                                                            'instantiate_custom_exceptions': True,
                                                            'import_custom_exceptions': True})
    print(repr(ts.service))
    config = Configuration()
    ts.service.config(config)
    ts.start()
