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
    Acts like a band master ;)
    (Try to) Provides useful resources to others.
"""
import sys
import os
sys.path.append(os.path.abspath(os.path.dirname(__file__)))
sys.executable = os.path.join(sys.exec_prefix, 'pythonw.exe')
import rpyc

from transitioncore.eventslistener.kernelcomeventslistener import KernelComEventListener
from transitioncore.eventslistener.kernelconfigeventlistener import KernelConfigurationEventListener
from transitioncore.exceptions.kernelexception import KernelException
from transitioncore.configuration import Configuration
from transitioncore.appmanager import AppManager


class TransitionKernel():
    """
    Acts like a band master ;)
    (Try to) Provides useful resources to others.
    """

    # TODO provide appmanager an excel event

    def __init__(self):
        """
        Kernel Init
        """
        self._com_app = None
        self._com_app_type = None
        self._addin = None

        #self._config = Configuration()

        # TODO : learn more... idea seems clear but I make something wrong.
        # get config like ConfigService is already launched
        try:
            self._conn = rpyc.connect("localhost", port=22)
            self._config = self._conn.root
        except ConnectionRefusedError:
            # Not launched, must do it
            from subprocess import Popen
            config_service_path = "{}{}".format(os.path.abspath(os.path.dirname(__file__)), "\\configservice.py")
            print("TransitionKernel : launch", config_service_path)
            Popen([sys.executable, config_service_path])
            self._config = rpyc.connect("localhost", port=22).root

        # register to configuration events. See KernelConfigurationEventListener
        # for further info on events and kernel actions.
        # self._config.add_event_listener(KernelConfigurationEventListener(self))

        # KernelComEventsListener instantiation
        self._com_events_listener = KernelComEventListener(self)

        # App Manager
        self._app_manager = None

    @property
    def config(self):
        """Get Configuration instance"""
        return self._config

    def set_com_app(self, com_app):
        """Set COM App and determine its type"""
        # What kind of COM Application it is ?
        app_description = repr(com_app)

        #looks for known com app
        for app_type in self._config.com_apps:
            if app_type.lower() in app_description:
                self._com_app_type = app_type

        if self._com_app_type is None:
            raise KernelException("TransitionKernel : {} is not handled".format(app_description))

        self._com_app = com_app

    def get_com_app(self):
        """Get COM App type"""
        if self._com_app is None:
            raise KernelException("TransitionKernel : application is not defined")
        return self._com_app

    com_app = property(get_com_app, set_com_app)

    @property
    def com_app_type(self):
        """
        Get COM app type
        :return: str. COM App type (Excel, ...)
        """
        if self._com_app is None:
            raise KernelException("TransitionKernel : application description is not defined")
        return self._com_app_type

    def set_addin(self, addin):
        self._addin = addin

    def get_addin(self):
        if self._addin is None:
            raise KernelException("TransitionKernel : addin is not defined")
        return self._addin

    addin = property(get_addin, set_addin)

    @property
    def kernel_com_events_listener(self):
        """
        Return COM Event Listener instance.
        :rtype : KernelComEventListener
        :return:  KernelComEventListener instance
        """
        return self._com_events_listener

    def run(self):
        """
        Kernel startup.
        :return: None
        """
        print('TransitionKernel launched !')

        for app_type in TransitionAppType:
            print('TransitionKernel enabled {} : {}'.format(app_type.value, ' '.join(
                x for x in self._config.get_enabled_app_list(app_type) if x)))

        if self._com_app.EnableEvents is False:
            print("TransitionKernel : Enabling Events !")
            self._com_app.EnableEvents = True

        self._app_manager = AppManager(self)
        self._app_manager.run()

    def terminate(self):
        """
        Kernel shutdown
        :return: None
        """
        print("TransitionKernel is terminating...")
        self._app_manager.terminate()
        print("TransitionKernel terminated")

    @staticmethod
    def transition_register(klass, com_app="Excel"):
        """
        Register add-in class in Excel
        :param klass:
        :return: None
        """

        # TODO make it register COM Add-in for wider range of Office Apps

        import win32com.server.register
        import winreg

        win32com.server.register.UseCommandLine(klass)

        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Office\\" + com_app + "\\Addins")
        subkey = winreg.CreateKey(key, klass._reg_progid_)
        winreg.SetValueEx(subkey, "CommandLineSafe", 0, winreg.REG_DWORD, 0)
        winreg.SetValueEx(subkey, "LoadBehavior", 0, winreg.REG_DWORD, 3)
        winreg.SetValueEx(subkey, "Description", 0, winreg.REG_SZ,
                          klass.addin_name)
        winreg.SetValueEx(subkey, "FriendlyName", 0, winreg.REG_SZ,
                          klass.addin_description)

    @staticmethod
    def transition_unregister(klass, com_app="Excel"):
        """
        Unregister add-in class (Excel)
        :param klass:
        :return: None
        """

        # TODO make it unregister COM Add-in for wider range of Office Apps

        import win32com.server.register
        import winreg

        win32com.server.register.UseCommandLine(klass)

        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER,
                             "Software\\Microsoft\\Office\\" + com_app + "\\Addins\\" + klass._reg_progid_)
        except WindowsError:
            pass