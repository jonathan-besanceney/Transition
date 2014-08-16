# ------------------------------------------------------------------------------
# Name:        eventsinterface.py
# Purpose:     COM Add-in Events Interface. You may want subclass it to :
#              - register an event listener on COM Add-in Events triggered by
#                the app.
#              - have a PEP 8 COM Event Interface to use.
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

class COMEventsInterface:
    """
    COM Add-in Events Interface. You may want subclass it to :
            - register an event listener on COM Add-in Events triggered by
                the app.
            - have a PEP 8 COM Event Interface to use.
    """
    def __init__(self):
        #sub classes will replace event_list tuple with their event name (eg on_connection...) tuple
        #in order to prevent call non-implemented methods
        self.event_list = ()

    def on_connection(self, application, connect_mode, addin, custom):
        pass

    def on_disconnection(self, mode, custom):
        pass

    def on_addins_update(self, custom):
        pass

    def on_startup_complete(self, custom):
        pass

    def on_begin_shutdown(self, custom):
        pass