# ------------------------------------------------------------------------------
# Name:        configuration
# Purpose:     Configuration Class
#
# Author:      Jonathan Besanceney <jonathan.besanceney@gmail.com>
#
# Created:     01/06/14
# Copyright:   (c) 2014 Jonathan Besanceney
#
# This file is a part of Transition
#
# Transition is free software: you can redistribute it and/or modify
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
import inspect

import pickle
import os
from os.path import expanduser
import pkgutil
import sys

from transitioncore import TransitionAppType, transition_app_path

sys.path.append(os.path.abspath(os.path.dirname(__file__)))

from transitioncore.eventdispatcher import TransitionEventDispatcher
from transitioncore.eventsinterface.configeventinterface import ConfigEventsInterface
from transitioncore.exceptions.configurationexception import ConfigurationException

# TODO : store conf in sqlite


class Configuration(TransitionEventDispatcher):
    """ Deals with Transition configuration."""

    def __init__(self):
        super(Configuration, self).__init__(ConfigEventsInterface)

        #all add-ins and apps regardless their statuses
        self._app_available_list = {TransitionAppType.excel_wbapp: list(),
                                    TransitionAppType.excel_addin: list()}
        self._update_available_app_list(None, False)

        #enabled add-ins and apps
        self._app_enabled_list = {TransitionAppType.excel_wbapp: list(),
                                  TransitionAppType.excel_addin: list()}

        #disabled add-ins and apps
        self._app_disabled_list = {TransitionAppType.excel_wbapp: list(),
                                   TransitionAppType.excel_addin: list()}

        #events listeners
        self._event_listener_list = list()

        self.enabled_apps_file = expanduser("~") + "/.enabled_apps.list"

        self._read_conf()

    def _update_available_app_list(self, app_type=None, fire_event=True):
        """
        Updates available app list. Fire on_app_add(self, addin_name) or on_app_remove(self, addin_name)
        on change
        """

        if app_type is None:
            for app_type in TransitionAppType:
                self._update_available_app_list(app_type, fire_event)
        else:
            app_list = list()
            # get available sub-packages of excelapps watching for addintion
            for _, name, is_package in pkgutil.iter_modules(transition_app_path[app_type]):
                if is_package and name not in self._app_available_list[app_type]:
                    app_list.append(name)
                    self._app_available_list[app_type].append(name)
                    if fire_event:
                        self._fire_event("on_app_add", (app_type, name))

            # now we want to see deletion
            for name in self._app_available_list[app_type]:
                if name not in app_list:
                    self._app_available_list[app_type].remove(name)
                    if fire_event:
                        self._fire_event("on_app_remove", (app_type, name))

    def _update_disabled_app_list(self, app_type=None):
        """

        :param app_type: TransitionAppType Enum
        """
        if app_type is None:
            for app_type in TransitionAppType:
                self._update_disabled_app_list(app_type)
        else:
            try:
                # get available sub-packages of app_type
                for _, name, is_package in pkgutil.iter_modules(transition_app_path[app_type]):
                    if is_package and name not in self._app_enabled_list[app_type]:
                        self._app_disabled_list[app_type].append(name)
            except TypeError as te:
                print("TypeError", repr(self._app_enabled_list), repr(te))

    def _read_conf(self):
        # Creates apps conf file if not exists
        if not os.path.isfile(self.enabled_apps_file):
            self._write_conf()

        config_object = pickle.load(open(self.enabled_apps_file, "rb"))

        #check if it's the good format
        if isinstance(config_object, dict):
            self._app_enabled_list = config_object
        else:
            #well, no. rewrite an empty dict
            self._write_conf()

        self._update_disabled_app_list()

    def _write_conf(self):
        pickle.dump(self._app_enabled_list, open(self.enabled_apps_file, "wb"))
        self._update_disabled_app_list()



    def enable_app(self, app_type, app_name):
        """
        Enables app_name and save configuration.
        :param app_type: TransitionAppType Enum
        :param app_name: name of the excel_app to enable.
        """

        if app_name in self._app_available_list[app_type]:
            if app_name in self._app_enabled_list[app_type]:
                mesg = "Configuration.enable_app() : app {} is already enabled !".format(app_name)
                print(mesg)
                raise ConfigurationException(mesg)
            else:
                print("Enabling {} app...".format(app_name))
                self._app_enabled_list[app_type].append(app_name)
                self._write_conf()
                self._update_disabled_app_list(app_type)

                #fire on_addin_enable event
                self._fire_event("on_app_enable", (app_type, app_name))
        else:
            mesg = "Configuration.enable_app() : Can't enable unavailable {} app !\nAvailable apps are {}.".format(
                app_name, repr(self._app_available_list[app_type]))
            print(mesg)
            raise ConfigurationException(mesg)



    def disable_app(self, app_type, app_name):
        """
        Disables app_name and save configuration
        :param app_type: TransitionAppType Enum
        :param app_name: name of the excel_app to disable
        """
        if app_name in self._app_enabled_list[app_type]:
            print("Disabling {} app...".format(app_name))
            self._app_enabled_list[app_type].remove(app_name)
            self._write_conf()
            self._update_disabled_app_list(app_type)

            #fire disable_app event
            self._fire_event("on_app_disable", (app_type, app_name))
        else:
            mesg = "Configuration.disable_app() : Can't disable unregistered {} app !".format(app_name)
            print(mesg)
            raise ConfigurationException(mesg)

    def get_disabled_app_list(self, app_type):
        """
        Returns disabled app list for given app_type
        :param app_type: TransitionAppType
        :return: disabled app list
        """
        return self._app_disabled_list[app_type]

    def get_enabled_app_list(self, app_type):
        """
        Returns enabled app list for given app_type
        :param app_type: TransitionAppType
        :return: enabled app list
        """
        return self._app_enabled_list[app_type]

    def get_app_list(self, app_type=None):
        """
        Return a list of tuple [(app_name, bool_status)] of available apps
        :rtype list
        """
        app_list = list()

        if app_type is None:
            for app_type in TransitionAppType:
                app_list.extend(self.get_app_list(app_type))
        else:
            for _, name, is_package in pkgutil.iter_modules(transition_app_path[app_type]):
                if is_package and name in self._app_enabled_list[app_type]:
                    app_list.append((name, True))
                elif is_package:
                    app_list.append((name, False))

        return app_list

    def get_app_status(self, app_type, app_name):
        """
        Returns status of given module

        :param app_name
        :return boolean as status
        """
        status = False
        app_list = self.get_app_list(app_type)
        for name, status in app_list:
            if name == app_name:
                break

        return status

    def print_app_list(self, app_type):
        """
        Displays names, state (loaded or not) and descriptions of the available
        excelapps.

        """

        print("Available {} :".format(app_type.value))

        app_list = self.get_app_list(app_type)

        for name, status in app_list:
            if status:
                print("\n* {} [ENABLED] : \n".format(name))
            else:
                print("\n* {} [DISABLED] : \n".format(name))

            print(self.get_app_desc(app_type, name))

    @staticmethod
    def get_app_desc(app_type, app_name):
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
            print("TransitionKernel.get_app_desc({}, {}) : ".format(app_type.value, app_name), repr(e))
            return -1