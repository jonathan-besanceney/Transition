# ------------------------------------------------------------------------------
# Name:         excelapphandler
# Purpose:      Handles Workbook Activate event to launch appropriate
#               application. Note this handler *will never close* Workbook
#               application for you. Closing your apps is your responsibility.
#
# Author:       Jonathan Besanceney <jonathan.besanceney@gmail.com>
#
# Created:      12/03/2014
# Copyright:    (c) 2014 Jonathan Besanceney
#
#    This file is a part of Transition
#
#    Transition is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Lesser General Public License as published by
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

# TODO : refactor this to manage all app_type. Events must move.
"""
 Handles Workbook Activate event to launch appropriate Workbook application.
 Note this handler *will kill* your Workbook applications at Excel shutdown.
 Closing your apps in a clean way is under your responsibility.
"""

from win32com.client import DispatchWithEvents

from transitioncore import TransitionAppType
from transitioncore.comeventslistener.amappeventlistener import AppManagerExcelEventListener


class AppManager():
    """Excel listener thread.
    Launched on TransitionMain.OnConnection()
    """

    com_app_event_class = {"excel": AppManagerExcelEventListener, }

    def __init__(self, com_app, config):
        """
        Set up handler thread
        :param com_app: COM Application instance
        """
        self.com_app = com_app
        self.com_app_type = None
        self.com_app_events = None
        self.config = config
        self.app_list = list()
        self.addin_list = list()

        # What kind of COM Application it is ?
        app_description = repr(self.com_app)
        if app_description.find("excel"):
            self.com_app_type = "excel"
            self.com_app_events = self.com_app_event_class["excel"]

    def run_app(self, app_type, app_name):
        #
        pass

    def terminate_app(self, app_type, app_name):
        pass

    def run(self):
        if self.com_app_events is not None:
            #launch COM App Addins

            #See for already opened documents. Are they handled ?

            #Start looking for COM App events
            self.com_app_events = DispatchWithEvents(self.com_app, self.com_app_events)
            self.com_app_events.set_app_manager(self)
            print("addin_list", repr(self.addin_list))

    def terminate(self):
        pass

