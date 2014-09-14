# ------------------------------------------------------------------------------
# Name:         excelapphandler
# Purpose:      Handles Workbook Activate event to launch appropriate
# application. Note this handler *will never close* Workbook
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
 AppManager :
 * Launches addin for current COM Application
 * Handles Workbook Activate event to launch appropriate document (eg. Workbook) application.

 Note AppManager *will kill* your Workbook applications at COM Application shutdown.
 Closing your apps in a clean way is under your responsibility (use terminate()).
"""

import inspect

from win32com.client import DispatchWithEvents
from transitioncore.configuration import Configuration

from transitioncore.comeventslistener.amappeventlistener import AppManagerExcelEventListener

class AppManager():
    """Excel listener thread.
    Launched on TransitionCOMEventsListener.OnConnection()
    """

    com_app_event_class = {"excel": AppManagerExcelEventListener, }

    def __init__(self, kernel):
        """
        Set up Application Manager.
        :param kernel: Kernel instance providing infos we need here.
        """
        self._kernel = kernel

        self.app_list = {"app": list(), "addin": list()}
        self.app_path = {"app": None, "addin": None}
        self.started_app_dict = {}

        if self._kernel.com_app_type is not None:
            # Retrieve stuff relative to COM Application
            self.com_app_events = self.com_app_event_class[self._kernel.com_app_type]

            self.app_list["app"] = self._kernel.config.get_enabled_app_list(
                self._kernel.config.transition_app_type_tree[self._kernel.com_app_type]["app"])

            self.app_list["addin"] = self._kernel.config.get_enabled_app_list(
                self._kernel.config.transition_app_type_tree[self._kernel.com_app_type]["addin"])

    def run_app(self, app_type, app_name, document=None):
        """
        Import, Instantiate, call run() and register app
        :param app_type: TransitionAppType
        :param app_name: Application Name
        :param document: opened document (eg workbook) to link with a document app
        """
        #TODO make it works for document apps
        #TODO generate pycache before launching app. https://docs.python.org/3/library/compileall.html#module-compileall

        print("AppManager.run_app({}, {})".format(app_type.value, app_name))
        # Import app module
        try:
            excel_addin_module = inspect.importlib.import_module("{}.{}".format(app_type.value, app_name))

            if document is None:
                # Init addin
                app = excel_addin_module.app_class(self._kernel.com_app)
            else:
                # Init app
                app = excel_addin_module.app_class(self._kernel.com_app, document)

            # Start it
            app.run()

            # Register it
            if app_type not in self.started_app_dict.keys():
                self.started_app_dict[app_type] = dict()

            if app_name not in self.started_app_dict[app_type].keys():
                self.started_app_dict[app_type][app_name] = app
                print("AppManager : launched", app_type, repr(self.started_app_dict[app_type].keys()))
            else:
                #TODO Raise a kind of "already launched" exception
                pass
        except TypeError as te:
            #TODO Raise a kind of "unknown app" exception
            pass

    def terminate_app(self, app_type, app_name, document=None):
        """
        call terminate() and unregister app
        :param app_type: TransitionAppType
        :param app_name: Application Name
        :param document: closing document (eg workbook)
        """

        if app_type in self.started_app_dict.keys():
            if app_name in self.started_app_dict[app_type].keys():
                self.started_app_dict[app_type][app_name].terminate()
                del self.started_app_dict[app_type][app_name]
                print("AppManager : launched", app_type, repr(self.started_app_dict[app_type].keys()))
            else:
                #TODO Raise a kind of "unknown app_name" exception
                pass
        else:
            #TODO Raise a kind of "unknown app_type" exception
            pass

    def run(self):
        if self.com_app_events is not None:
            #launch COM App Addins
            for addin in self.app_list["addin"]:
                self.run_app(self._kernel.config.transition_app_type_tree[self._kernel.com_app_type]["addin"], addin)

            #TODO See for already opened documents. Are they handled ?

            #Start looking for COM App events
            self.com_app_events = DispatchWithEvents(self._kernel.com_app, self.com_app_events)
            self.com_app_events.set_app_manager(self)
        else:
            #TODO Raise a kind of "not initialized" exception
            pass

    def terminate(self):
        import copy
        print("AppManager : terminating...")
        app_type = self._kernel.config.transition_app_type_tree[self._kernel.com_app_type]["addin"]
        if app_type in self.started_app_dict.keys():
            addin_dict = copy.copy(self.started_app_dict[app_type])
            for addin in addin_dict.keys():
                self.terminate_app(app_type, addin)

        print("AppManager : terminated.")

    # def launch_wb_app(self, wb) -> ExcelWorkbookAppSkell:
    #     """
    #     Launches an ExcelApp on the given wb.
    #     :param wb:
    #     :return: Thread instance if ok. None if no app are found.
    #     :rtype: ExcelWorkbookAppSkell
    #     """
    #     launch = True
    #     wb_app = None
    #     for t in self._current_app_list:
    #         # Checks if app already launched
    #         if hasattr(t, "wb") and t.wb.Name == wb.Name:
    #             # We don't want to launch it again
    #             launch = False
    #
    #     if launch:
    #         wb_app = get_wb_app_instance(wb)
    #
    #         if wb_app is not None:
    #             print("Launching Transition Workbook App {} on {} ...".format(wb_app.name, wb.Name))
    #             # wb_app.daemon = True
    #             wb_app.run()
    #
    #     return wb_app