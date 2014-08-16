# ------------------------------------------------------------------------------
# Name:         excelapphandler
# Purpose:      Handles Workbook Activate event to launch appropriate
# application. Note this handler *will never close* Workbook
# application for you. Closing your apps is your responsibility.
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


  - Excel Add-in register/unregister methods. (from <ekoome@yahoo.com> Eric Koome's
  /win32com/demo/excelAddin.py)
"""
import sys
import os
sys.path.append(os.path.abspath(os.path.dirname(__file__)))

import inspect
from win32com.client import DispatchWithEvents

from excelapps import get_wb_app_instance
from excelapps.appskell import ExcelWorkbookAppSkell
from transitioncore import defaultNamedNotOptArg
from transitioncore.comeventsinterface.excelappeventsinterface import ExcelAppEventsInterface
from transitioncore.eventslistener.kernelcomeventslistener import KernelComEventListener
from transitioncore.eventslistener.kernelconfigeventlistener import KernelConfigurationEventListener
from transitioncore.exceptions.kernelexception import KernelException
from transitioncore.configuration import Configuration

from transitioncore import TransitionAppType


class TransitionKernel():
    """

    """

    def __init__(self, waitExcelVisible=False):
        """
        Set up handler thread
        :param excel_app: Excel Application instance
        :param waitExcelVisible: tells to the handler to wait until Excel is visible before launch
        """
        self._application = None
        self._addin = None
        self._waitExcelVisible = waitExcelVisible

        self._config = Configuration()
        # register to configuration events
        self._config.add_event_listener(KernelConfigurationEventListener(self))

        # KernelComEventsListener instantiation
        self._com_events_listener = KernelComEventListener(self)

    def set_application(self, application):
        self._application = application

    def get_application(self):
        if self._application is None:
            raise KernelException("TransitionKernel : application is not defined")
        return self._application

    def set_addin(self, addin):
        self._addin = addin

    def get_addin(self):
        if self._addin is None:
            raise KernelException("TransitionKernel : addin is not defined")
        return self._addin

    def get_kernel_com_events_listener(self):
        """
        :rtype : KernelComEventListener
        :return:  KernelComEventListener instance
        """
        return self._com_events_listener

    def run(self):
        print('TransitionKernel launched !')

        for app_type in TransitionAppType:
            print('TransitionKernel enabled {} : {}'.format(app_type.value, ' '.join(
                x for x in self._config.app_get_enabled_list(app_type) if x)))

        if self._waitExcelVisible is False and self._application.Visible == 0:
            print("TransitionKernel : Excel wasn't running... Exiting...")
            self._application.Quit()
        else:
            if self._application.EnableEvents is False:
                print("TransitionKernel : Enabling Events !")
                self._application.EnableEvents = True

            # TODO : move com event handling to app_manager. kernel is a bandmaster :)
            self._application = DispatchWithEvents(self._application, TransitionEvents)
            self._application.name = "TransitionKernel ExcelEvent"

    def terminate(self):
        print("TransitionKernel is terminating...")
        self._application = None
        print("TransitionKernel terminated...")

    def launch_wb_app(self, wb) -> ExcelWorkbookAppSkell:
        """
        Launches an ExcelApp on the given wb.
        :param wb:
        :return: Thread instance if ok. None if no app are found.
        :rtype: ExcelWorkbookAppSkell
        """
        launch = True
        wb_app = None
        for t in self._current_app_list:
            # Checks if app already launched
            if hasattr(t, "wb") and t.wb.Name == wb.Name:
                # We don't want to launch it again
                launch = False

        if launch:
            wb_app = get_wb_app_instance(wb)

            if wb_app is not None:
                print("Launching Transition Workbook App {} on {} ...".format(wb_app.name, wb.Name))
                # wb_app.daemon = True
                wb_app.run()

        return wb_app

    @staticmethod
    def app_get_desc(app_type, app_name):
        """
        Return the description of the given app

        :param app_name
        :rtype str
        :returns module description
        """

        try:
            # Dynamic import of the package - to be able to load comments
            module = inspect.importlib.import_module("{}.{}".format(app_type.value, app_name))
            # return top comments of the package
            return inspect.getcomments(module)

        except Exception as e:
            print("TransitionKernel.app_get_desc({}, {}) : ".format(app_type.value, app_name), repr(e))
            return -1

    @staticmethod
    def transition_register(klass):
        import win32com.server.register
        import winreg

        win32com.server.register.UseCommandLine(klass)

        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Office\\Excel\\Addins")
        subkey = winreg.CreateKey(key, klass._reg_progid_)
        winreg.SetValueEx(subkey, "CommandLineSafe", 0, winreg.REG_DWORD, 0)
        winreg.SetValueEx(subkey, "LoadBehavior", 0, winreg.REG_DWORD, 3)
        winreg.SetValueEx(subkey, "Description", 0, winreg.REG_SZ,
                          klass.addin_name)
        winreg.SetValueEx(subkey, "FriendlyName", 0, winreg.REG_SZ,
                          klass.addin_description)

    @staticmethod
    def transition_unregister(klass):
        import win32com.server.register
        import winreg

        win32com.server.register.UseCommandLine(klass)

        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER,
                             "Software\\Microsoft\\Office\\Excel\\Addins\\" + klass._reg_progid_)
        except WindowsError:
            pass


class TransitionEvents(ExcelAppEventsInterface):
    """
    This event class is used in TransitionKernel for launching appropriate workbook application.
    """

    def __init__(self):
        super(TransitionEvents, self).__init__()
        self.current_app_list = list()

    def add_event_handles(self, pyhandles):
        for pyhandle in pyhandles:
            self.add_event_handle(pyhandle)

    def add_event_handle(self, pyhandle):
        """This method will refer to TransitionKernel.add_event_handle"""
        pass

    def launch_wb_app(self, Wb) -> ExcelWorkbookAppSkell:
        pass

    def OnWindowActivate(self, Wb=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
        """
        OnWindowActivate is responsible for launching the right workbook handler.
        IMPORTANT : Workbook apps are responsible for their own termination.
        :param Wb: Workbook object
        :param Wn: Window object
        """
        print("{} OnWindowActivate {} {}".format(self.name, Wb.Name, Wn.Caption))
        wb_thread = self.launch_wb_app(Wb)
        if wb_thread is not None:
            self.current_app_list.append(wb_thread)
            self.add_event_handles(wb_thread.get_event_handles())


