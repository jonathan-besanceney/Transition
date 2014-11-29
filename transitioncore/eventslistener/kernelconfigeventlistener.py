# ------------------------------------------------------------------------------
# Name:        kernelconfigeventlistener
# Purpose:     ConfigEventsInterface implementation for Kernel usage
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

from transitioncore.eventsinterface.configeventinterface import ConfigEventsInterface
from rpyc import Service


class KernelConfigurationEventListener(ConfigEventsInterface):
    def __init__(self, kernel=None):
        super(KernelConfigurationEventListener, self).__init__()
        self.kernel = kernel
        #declare interesting kernel event. on_app_add and on_app_del are not kernel
        #stuff by now (more configuration GUI oriented)
        self.event_list = ("on_app_enable", "on_app_disable")

    def on_app_disable(self, app_type, app_name, com_app_tuple):
        print("KernelConfigurationEventListener.on_app_disable", app_type, app_name, com_app_tuple)

    def on_app_enable(self, app_type, app_name, com_app_tuple):
        print("KernelConfigurationEventListener.on_app_enable", app_type, app_name, com_app_tuple)