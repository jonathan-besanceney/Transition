# ------------------------------------------------------------------------------
# Name:        transition.py
# Purpose:     Transition Excel/COM Add-in launches an excel handler.
# Handler watches for documents open and close in a separate thread in order to
# launch appropriate excelapps to handle the workbooks automation.
# Registered exceladdins are launched at startup.
#
# Author:      Jonathan Besanceney <jonathan.besanceney@gmail.com> from
#               <ekoome@yahoo.com> Eric Koome's /win32com/demo/excelAddin.py
#
# Created:     14/03/2014
# Copyright:   (c) 2014 Jonathan Besanceney
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
Transition Excel/COM Add-in launches an excel handler.
Handler watches for documents open and close in a separate thread in order to
launch appropriate excelapps to handle the workbooks automation.
Registered exceladdins are launched at startup.
"""
import sys
# specify free threading, common way to think threading.
sys.coinit_flags = 0

import os

sys.path.append(os.path.abspath(os.path.dirname(__file__)))

import threading

from win32com import universal
from win32com.client import gencache, Dispatch
# Last import, at least after sys.coinit_flags = 0 to initialize it in
# free threading.
import pythoncom
import win32trace

import exceladdins
import transitionconfig
from transitioncore import transition_register, transition_unregister
from transitioncore.excelapphandler import WorkbookAppHandler

# Support for COM objects we use.
# Excel 2010
excel_application = gencache.EnsureModule('{00020813-0000-0000-C000-000000000046}', 0, 1, 7,
                                          bForDemand=True)

# Office 2010
gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 5,
                      bForDemand=True)

# The TLB defining the interfaces we implement in
universal.RegisterInterfaces('{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}', 0, 1, 0,
                             ["_IDTExtensibility2"])


class TransitionMain:
    """
    Transition main class.

    This class is registered in the Windows Registry to tell Excel to open
    pythoncom.dll to exec our program.

    """

    _com_interfaces_ = ['_IDTExtensibility2']
    _public_methods_ = []
    _reg_clsctx_ = pythoncom.CLSCTX_INPROC_SERVER
    _reg_clsid_ = "{C5482ECA-F559-45A0-B078-B2036E6F011A}"
    _reg_progid_ = "Python.Transition.ExcelAddin"
    _reg_policy_spec_ = "win32com.server.policy.EventHandlerPolicy"
    addin_name = "Transition Excel/COM Addin"
    addin_description = "Transition Excel/COM Addin enables you automate Excel with Python"

    def __init__(self):
        self.name = threading.currentThread().getName()
        print(self.addin_name, ": init", self.name)
        self.appHostApp = None
        self.excel_handler = None
        self.addin_thread_list = []

    def OnConnection(self, application, connect_mode, addin, custom):
        """
        Addin startup
        """
        win32trace.InitRead()
        self.appHostApp = application

        print(self.addin_name, "({}) : OnConnection".format(self.name), application,
              connect_mode, addin, custom)

        self.excel_handler = WorkbookAppHandler(self.appHostApp, True)

    def OnDisconnection(self, mode, custom):
        """
        Addin shutdown
        """
        print(self.addin_name, "({}) : OnDisconnection".format(self.name), mode, custom)

        print(self.addin_name, "({}) : shutting down addins threads...".format(self.name))
        exceladdins.unload_addins(self.addin_thread_list)
        print(self.addin_name, "({}) : addin threads are terminated.".format(self.name))

        print(self.addin_name, "({}) : shutting down handler thread...".format(self.name))
        self.excel_handler.quit()
        if self.excel_handler.is_alive():
            self.excel_handler.join()

        print(self.addin_name, "({}) : handler thread is terminated.".format(self.name))
        self.appHostApp.Quit()
        self.appHostApp = None
        win32trace.TermRead()

    def OnAddInsUpdate(self, custom):
        print(self.addin_name, "({}) : OnAddInsUpdate".format(self.name), custom)

    def OnStartupComplete(self, custom):
        # While Excel finish its startup process
        print(self.addin_name, "({}) : OnStartupComplete".format(self.name), custom)
        self.excel_handler.start()
        print(self.addin_name, "({}) : handler thread launched !".format(self.name))

        print(self.addin_name, "({}) : loading python addins...".format(self.name))
        self.addin_thread_list = exceladdins.load_addins(self.appHostApp)
        print(self.addin_name, "({}) : addins loaded !".format(self.name))

    def OnBeginShutdown(self, custom):
        # Excel begins to launch its termination process.
        print(self.addin_name, "({}) : OnBeginShutdown".format(self.name), custom)


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(
        description=
        "Transition Excel Add-in launchs an excel handler watching for documents open\n"
        + "and close in a separate thread.\n"
        + "Handler tries to launch appropriate excelapps to handle the workbooks.\n"
        + "Registered exceladdins are launched at startup.\n"
        + "Running transition.py without parameters registers add-in. See options bellow.")

    group = parser.add_mutually_exclusive_group()

    group.add_argument("--debug", help="registers Transition Excel Add-in in debug mode.\n"
                                       + "This option enables execution traces to be collected by the config add-in.",
                       action="store_true")
    group.add_argument("--unregister", help="unregisters register Transition Excel Add-in in debug mode.",
                       action="store_true")

    group.add_argument("-l", "--list", help="lists available excel pps and add-ins", action="store_true")

    group.add_argument("-al", "--app-list", help="lists available excelapps", action="store_true")
    group.add_argument("-ae", "--app-enable", help="enables available excel_app", type=str,
                       choices=transitionconfig.app_get_disabled_list())
    group.add_argument("-ad", "--app-disable", help="disables previously enabled excel_app", type=str,
                       choices=transitionconfig.app_get_enabled_list())

    group.add_argument("-dl", "--addin-list", help="lists available exceladdins", action="store_true")
    group.add_argument("-de", "--addin-enable", help="enables available exceladdins", type=str,
                       choices=transitionconfig.addin_get_disabled_list())
    group.add_argument("-dd", "--addin-disable", help="disables previously enabled exceladdins", type=str,
                       choices=transitionconfig.addin_get_enabled_list())

    args = parser.parse_args()

    if args.list:
        transitionconfig.app_print_list()
        transitionconfig.addin_print_list()
    elif args.app_list:
        transitionconfig.app_print_list()
    elif args.addin_list:
        transitionconfig.addin_print_list()
    elif args.app_enable is str:
        transitionconfig.app_enable(args.app_enable)
    elif args.app_disable is str:
        transitionconfig.app_disable(args.app_disable)
    elif args.addin_enable is str:
        transitionconfig.app_enable(args.addin_enable)
    elif args.addin_disable is str:
        transitionconfig.app_disable(args.addin_disable)
    elif args.unregister:
        import win32com.server.register
        win32com.server.register.UseCommandLine(TransitionMain)
        transition_unregister(TransitionMain)
    else:
        import win32com.server.register
        win32com.server.register.UseCommandLine(TransitionMain)
        transition_register(TransitionMain)
