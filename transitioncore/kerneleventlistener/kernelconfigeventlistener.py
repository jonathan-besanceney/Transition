# ------------------------------------------------------------------------------
# Name:        Script Name 
# Purpose:     TODO 
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

from transitioncore.configeventinterface.configeventinterface import ConfigEventsInterface


class KernelConfigurationEventListener(ConfigEventsInterface):
    def __init__(self, kernel=None):
        self.kernel = kernel

    def on_addin_add(self, addin_name):
        pass

    def on_addin_disable(self, addin_name):
        pass

    def on_addin_enable(self, addin_name):
        pass

    def on_addin_remove(self, addin_name):
        pass

    def on_app_add(self, app_name):
        pass

    def on_app_disable(self, app_name):
        pass

    def on_app_enable(self, app_name):
        pass

    def on_app_remove(self, app_name):
        pass