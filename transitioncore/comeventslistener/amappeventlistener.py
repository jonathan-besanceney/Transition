# ------------------------------------------------------------------------------
# Name:        amappeventlistener
# Purpose:     Application Manager Application (COM) Event Listener
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
import win32event
from transitioncore import defaultNamedNotOptArg
from transitioncore.comeventsinterface.excelappeventsinterface import ExcelAppEventsInterface


class AppManagerExcelEventListener(ExcelAppEventsInterface):
    """
    This event listener class is used in AppManager
    """

    def __init__(self):
        super(AppManagerExcelEventListener, self).__init__()
        self.app_manager = None

    def set_app_manager(self, app_manager):
        self.app_manager = app_manager

    def OnWindowActivate(self, Wb=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
        """
        :param Wb: Workbook object
        :param Wn: Window object
        """
        pass

    def OnWorkbookDeactivate(self, Wb=defaultNamedNotOptArg):
        pass

    def OnWorkbookBeforeClose(self, Wb=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
        pass
