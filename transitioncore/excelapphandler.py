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
"""
 Handles Workbook Activate event to launch appropriate Workbook application.
 Note this handler *will kill* your Workbook applications at Excel shutdown.
 Closing your apps in a clean way is under your responsibility.
"""
import sys
# specify free threading, common way to think threading.
sys.coinit_flags = 0

from threading import Thread
import threading

import win32event
import pythoncom
from win32com.client import DispatchWithEvents

from excelapps import launch_wb_app
from transitioncore import WAIT_FOR_EVENT_MSEC, defaultNamedNotOptArg
from transitioncore.excelappevents import ExcelAppEvents
import transitionconfig


class WorkbookAppHandler(Thread):
    """Excel listener thread.
    Launched on TransitionMain.OnConnection()
    """

    def __init__(self, xlApp, waitExcelVisible=False):
        """
        Set up handler thread
        :param xlApp: Excel Application instance
        :param waitExcelVisible: tells to the handler to wait until Excel is visible before launch
        """
        super(WorkbookAppHandler, self).__init__()
        self.xlApp = xlApp
        self._ask_quit = False
        self._waitExcelVisible = waitExcelVisible
        self.thread_name = threading.currentThread().getName()

    def quit(self):
        """Ask for thread termination"""
        self._ask_quit = True

    def run(self):
        print('WorkbookAppHandler {} launched !'.format(self.name))
        print('WorkbookAppHandler {} enabled apps : {}'
              .format(self.name, ' '.join(x for x in transitionconfig.app_get_enabled_list() if x)))
        try:
            if self._waitExcelVisible is False and self.xlApp.Visible == 0:
                print("WorkbookAppHandler {} : Excel wasn't running... Exiting..."
                      .format(self.name))
                self.xlApp.Quit()
            else:
                if self.xlApp.EnableEvents is False:
                    print("WorkbookAppHandler {} : Enabling Events !"
                          .format(self.name))

                self.xlApp = DispatchWithEvents(self.xlApp, WorkbookHandlerEvents)
                self.xlApp.name = "WorkbookAppHandler ExcelEvent"

                #open workbook apps for already opened workbooks
                for wb in self.xlApp.Workbooks:
                    launch_wb_app(wb)

                #Main loop. Will stop at excel termination. See TransitionMain.OnDisconnection
                while self._ask_quit is False:
                    win32event.WaitForSingleObject(self.xlApp.event, WAIT_FOR_EVENT_MSEC)

                #kill opened workbook apps
                #TODO

            print("WorkbookAppHandler {} is terminating...".format(self.name))

        except KeyboardInterrupt:
            print("WorkbookAppHandler {} : interruption exception".format(self.name),
                  "intercepted (termination asked) !")
        except pythoncom.com_error as details:
            print("WorkbookAppHandler {} Exception (com_error) : {}".format(self.name, details))
        except Exception as e:
            print("WorkbookAppHandler {} Exception".format(self.name, e))

        self.xlApp = None
        print("WorkbookAppHandler {} terminated...".format(self.name))


class WorkbookHandlerEvents(ExcelAppEvents):
    """
    This event class is used in WorkbookAppHandler for launching appropriate workbook application.
    """

    def __init__(self):
        super(WorkbookHandlerEvents, self).__init__()

    def OnWindowActivate(self, Wb=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
        """
        OnWindowActivate is responsible for launching the right workbook handler.
        IMPORTANT : Workbook apps are responsible for their own termination.
        :param Wb: Workbook object
        :param Wn: Window object
        """
        print("{} OnWindowActivate {} {}".format(self.name, Wb.Name, Wn.Caption))
        launch_wb_app(Wb)

        win32event.SetEvent(self.event)