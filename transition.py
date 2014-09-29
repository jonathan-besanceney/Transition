# ------------------------------------------------------------------------------
# Name:        transition.py
# Purpose:     Transition Office/COM Add-in.
#              Implements COM Event Interface to start and stop Transition
#              Kernel.
#              See transition.py --help for more details.
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
Transition Office/COM Add-in aims are to :
- Bootstrap kernel via TransitionCOMEventsListener when this class is registered as
  COM Add-in in Office Software (only Excel by now)
- Provide a command line interface to configuration. transition.py --help for more
  information.
"""

import sys
import os
sys.coinit_flags = 0

from win32com import universal
from win32com.client import gencache
import pythoncom
import win32trace

sys.path.append(os.path.abspath(os.path.dirname(__file__)))

from transitioncore.eventdispatcher import TransitionEventDispatcher
from transitioncore.kernel import TransitionKernel
from transitioncore.eventsinterface.comeventsinterface import COMEventsInterface

# Support for COM objects we use.
# Excel 2010
excel_application = gencache.EnsureModule('{00020813-0000-0000-C000-000000000046}', 0, 1, 7,
                                          bForDemand=True)
# TODO : EnsureModule for all office apps 'Access', 'MS Project', 'OneNote', 'Outlook', 'PowerPoint', 'Word'

# Office 2010
gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 5,
                      bForDemand=True)

# The TLB defining the interfaces we implement in
universal.RegisterInterfaces('{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}', 0, 1, 0,
                             ["_IDTExtensibility2"])


class TransitionCOMEventsListener(TransitionEventDispatcher):
    """
    Transition COM Add-In Event Class. Add-In entry point.

    Bootstrap Transition Kernel.

    This class is registered in the Windows Registry to tell Office app to open
    pythoncom.dll at startup to instantiate this class.

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
        super(TransitionCOMEventsListener, self).__init__(COMEventsInterface)
        self.name = TransitionCOMEventsListener.addin_name
        print(self.addin_name, ": init", self.name)
        self.com_app = None

        # Bootstrap kernel
        self._kernel = TransitionKernel()

        # reverse-register kernel com events listener
        self.add_event_listener(self._kernel.kernel_com_events_listener)

    def OnConnection(self, application, connect_mode, addin, custom):
        """
        Addin startup
        """
        print(self.addin_name, ": OnConnection", application, connect_mode, addin, custom)

        # link application to our add-in
        self.com_app = application

        # fire registered events
        self._fire_event("on_connection", (application, connect_mode, addin, custom))

    def OnDisconnection(self, mode, custom):
        """
        Addin shutdown
        """
        print(self.addin_name, ": OnDisconnection", mode, custom)

        # fire registered events from registered listeners
        self._fire_event("on_disconnection", (mode, custom))

        self.com_app.Quit()
        self.com_app = None
        win32trace.TermRead()

    def OnAddInsUpdate(self, custom):
        print(self.addin_name, ": OnAddInsUpdate", custom)

        # fire registered events
        self._fire_event("on_addins_update", custom)

    def OnStartupComplete(self, custom):
        # While Excel finish its startup process
        print(self.addin_name, ": OnStartupComplete", custom)

        # fire registered events
        self._fire_event("on_startup_complete", custom)

    def OnBeginShutdown(self, custom):
        # Excel begins to launch its termination process.
        print(self.addin_name, ": OnBeginShutdown", custom)

        # fire registered events
        self._fire_event("on_begin_shutdown", custom)


if __name__ == '__main__':
    import argparse
    from transitioncore.configuration import Configuration, ConfigurationException

    try:
        config = Configuration()

        parser = argparse.ArgumentParser(
            description=
            "Transition Office Add-in launches a handler watching for documents open\n"
            + "and close.\n"
            + "Handler tries to launch appropriate docapp to handle opened office document.\n"
            + "Registered addin are launched at startup.\n"
            + "Running transition.py without parameters registers add-in.\n"
            + "See optional arguments bellow :")

        group = parser.add_mutually_exclusive_group()

        # TODO : cleanup arg list (like Configuration interface)

        group.add_argument("--debug", help="registers Transition Excel Add-in in debug mode.\n"
                                           + "This option enables execution traces to be collected by the config add-in.",
                           action="store_true")
        group.add_argument("--unregister", help="unregisters register Transition Excel Add-in in debug mode.",
                           action="store_true")

        parser.add_argument("list", help="lists available doc-apps and add-ins. Optional arg {}"
                            .format(Configuration.app_type_list), action="store_true")
        parser.add_argument("enable", help="enables available doc-apps and add-ins", action="store_true")
        parser.add_argument("disable", help="disables previously enabled doc-apps and add-ins", action="store_true")
        parser.add_argument("args", nargs=argparse.REMAINDER, type=str)
        args = parser.parse_args()

        if args.list:
            if len(args.args) == 2:
                config.print_app_list(args.args[1])
            elif len(args.args) > 2:
                print("list command accept one optional application type argument {}. Command line args : {}."
                      .format(Configuration.app_type_list, args.args))
            else:
                for app_type in Configuration.app_type_list:
                    config.print_app_list(app_type)
        elif args.enable:
            config.enable_app('docapp', args.enable_app)
        elif args.disable:
            config.disable_app('docapp', args.disable_app)
        elif args.unregister:
            TransitionKernel.transition_unregister(TransitionCOMEventsListener)
        else:
            TransitionKernel.transition_register(TransitionCOMEventsListener)
    except ConfigurationException as ce:
        print("Transition configuration command returned an error :", ce.value)