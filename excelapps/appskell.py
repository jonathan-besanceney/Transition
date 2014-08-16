# ------------------------------------------------------------------------------
# Name:        app_skell
# Purpose:     Define a standard way to replace VBA in Excel Workbooks by an
#              external COM application launched by Transition Excel/COM Add-in.
#
# Author:       Jonathan Besanceney <jonathan.besanceney@gmail.com>
#
# Created:     09/03/2014
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
"""
Define a standard way to replace VBA in Excel Workbooks by an
external COM application launched by Transition Excel/COM Add-in.
"""


import sys

# specify free threading, common way to think threading.
sys.coinit_flags = 0

import os

sys.path.append(os.path.abspath(os.path.dirname(__file__)))
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir)))

from win32com.client import DispatchWithEvents
import xl as pyvot

from transitioncore.comeventsinterface.excelwbeventsinterface import ExcelWbEventsInterface


class ExcelWbEventsSkell(ExcelWbEventsInterface):
    """
    Workbook Event handling basics
    """

    def __init__(self):
        super(ExcelWbEventsSkell, self).__init__()
        self.before_close = False
        self.name = 'ExcelWbEventsSkell'

    def del_event_handle(self, pyhandle):
        pass

    def OnActivate(self):
        print(self.name, self.Name, "ExcelWbEventsSkell OnActivate")
        # UnSet the before_close flag to help OnDeactivate to determine
        # if we need quit
        self.before_close = False

    def OnBeforeClose(self, Cancel):
        print(self.name, self.Name, "ExcelWbEventsSkell OnBeforeClose", Cancel,
              repr(self.pywb))
        # Set the before_close flag to help OnDeactivate to determine
        # if we need quit
        self.before_close = True
        return Cancel

    def OnDeactivate(self):
        # here it's the Name property of the Excel Workbook
        print(self.name, self.Name, "ExcelWbEventsSkell OnDeactivate")
        # Is this event appends after OnBeforeClose ?
        if self.before_close:
            # We want quit, the workbook is closing now.
            self.del_event_handle(self.event)
            print(self.name, self.Name, "ExcelWbEventsSkell remove event handle")


class ExcelWorkbookAppSkell():
    """Application Standard Skeleton"""

    def __init__(self, wb, evt_handler, name):
        super(ExcelWorkbookAppSkell, self).__init__()
        print("{} : init on {}".format(name, wb.Name))
        self.wb = wb
        self.name = name
        self.evt_handler = evt_handler
        self._pyhandles = list()

    def add_event_handle(self, pyhandle):
        self._pyhandles.append(pyhandle)

    def del_event_handle(self, pyhandle):
        self._pyhandles.remove(pyhandle)

    def get_event_handles(self):
        return self._pyhandles

    @staticmethod
    def is_handled_workbook(wb):
        """
        This method returns an instance of ExcelWorkbookAppSkell if wb is a
        handled Workbook

        :param wb: Workbook instance
        :return ExcelWorkbookAppSkell instance
        """
        return ExcelWorkbookAppSkell(wb, ExcelWbEventsSkell, "ExcelWorkbookAppSkell")

    def run(self):
        """ Initialize and launch application main loop"""

        print("{} : Init Transition WorkbookApp on {}".format(self.name,  self.wb.Name))

        if self.wb is not None:
            # Add Events Handlers to the Workbook instance
            self.wb = DispatchWithEvents(self.wb, self.evt_handler)
            self.wb.name = self.name

            # Get the Pythonic interface to Excel from Microsoft
            # (Pyvot => xl)
            self.wb.pywb = pyvot.Workbook(self.wb)

            self.wb.del_event_handle = self.del_event_handle

            print("{} : Transition is plugged on {}. Waiting for events...".format(self.name, self.wb.Name))

        else:
            print("{} : You must give a valid Workbook instance !".format(self.name))

if __name__ == '__main__':
    from win32com.client import Dispatch
    xlApp = Dispatch("Excel.Application")
    xlApp.Visible = 1
    xlApp.EnableEvents = True
    m_wb = xlApp.Workbooks.Add()
    app = ExcelWorkbookAppSkell.is_handled_workbook(m_wb)
    app.run()
    app = None
    xlApp = None
    sys.exit(0)