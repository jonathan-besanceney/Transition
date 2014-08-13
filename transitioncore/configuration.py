# ------------------------------------------------------------------------------
# Name:        Script Name 
# Purpose:     TODO 
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

import exceladdins
import excelapps

from transitioncore.configeventinterface.configeventinterface import ConfigEventsInterface
from transitioncore.kernelexception.configurationexception import ConfigurationException


class Configuration():
    """ Deals with Transition configuration."""

    def __init__(self):
        #add-ins and apps sub-packages path
        self._pkg_addin_path = exceladdins.__path__
        self._pkg_app_path = excelapps.__path__

        #all add-ins and apps regardless their statuses
        self._addin_available_list = list()
        self.addin_update_available_list(False)
        self._app_available_list = list()
        self.app_update_available_list(False)

        #enabled add-ins and apps
        self._addin_enabled_list = list()
        self._app_enabled_list = list()

        #disabled add-ins and apps
        self._addin_disabled_list = list()
        self._app_disabled_list = list()

        #events listeners
        self._event_listener_list = list()

        self.enabled_addins_file = expanduser("~") + "/.enabled_addins.list"
        self.enabled_apps_file = expanduser("~") + "/.enabled_apps.list"

        # Creates apps conf file if not exists
        if not os.path.isfile(self.enabled_apps_file):
            self._write_apps_conf()

        self._read_apps_conf()

        # Creates add-ins conf file if not exists
        if not os.path.isfile(self.enabled_apps_file):
            self._write_addin_conf()

        self._read_addin_conf()

    def get_addin_pkg_path(self):
        """
        Returns add-in package path
        :return: path to add-in package
        """
        return self._pkg_addin_path

    def get_app_pkg_path(self):
        """
        Returns app package path
        :return: path to app package
        """
        return self._pkg_app_path

    def addin_update_available_list(self, fire_event=True):
        """
        Updates available add-in list. Fire on_addin_add(self, addin_name) or on_addin_remove(self, addin_name)
        on change
        """
        addin_list = list()
        # get available sub-packages of exceladdins watching for addintion
        for _, name, is_package in pkgutil.iter_modules(self._pkg_addin_path):
            if is_package and name not in self._addin_available_list:
                addin_list.append(name)
                self._addin_available_list.append(name)
                if fire_event:
                    for listener in self._event_listener_list:
                        listener.on_addin_add(name)

        # now we want to see deletion
        for name in self._addin_available_list:
            if name not in addin_list:
                self._addin_available_list.remove(name)
                if fire_event:
                    for listener in self._event_listener_list:
                        listener.on_addin_remove(name)

    def app_update_available_list(self, fire_event=True):
        """
        Updates available app list. Fire on_app_add(self, addin_name) or on_app_remove(self, addin_name)
        on change
        """
        app_list = list()
        # get available sub-packages of excelapps watching for addintion
        for _, name, is_package in pkgutil.iter_modules(self._pkg_app_path):
            if is_package and name not in self._app_available_list:
                app_list.append(name)
                self._app_available_list.append(name)
                if fire_event:
                    for listener in self._event_listener_list:
                        listener.on_app_add(name)

        # now we want to see deletion
        for name in self._app_available_list:
            if name not in app_list:
                self._app_available_list.remove(name)
                if fire_event:
                    for listener in self._event_listener_list:
                        listener.on_app_remove(name)

    def add_config_event_listener(self, listener):
        """
        Add a Config Event listener. Allow you to be notified of all configuration events.
        :param listener: ConfigEventsInterface instance
        :raise ConfigurationException:
        """
        if isinstance(listener, ConfigEventsInterface):
            self._event_listener_list.append(listener)
        else:
            mesg = "Configuration.add_config_event_listener() : Can't add non ConfigEventsInterface listener", \
                   repr(listener)
            print(mesg)
            raise ConfigurationException(mesg)

    def del_config_event_listener(self, listener):
        """
        Remove a previously added Config Event listener.
        :param listener: registered ConfigEventsInterface
        :raise ConfigurationException: raised if listener is not registered
        """
        try:
            self._event_listener_list.remove(listener)
        except ValueError:
            mesg = "Configuration.del_com_event_listener() : Can't remove unregistered listener", repr(listener)
            print(mesg)
            raise ConfigurationException(mesg)

    def _read_addin_conf(self):
        self._addin_enabled_list = pickle.load(open(self.enabled_addins_file, "rb"))
        self._update_addin_disabled_list()

    def _write_addin_conf(self):
        pickle.dump(self._addin_enabled_list, open(self.enabled_addins_file, "wb"))
        self._update_addin_disabled_list()

    def _read_apps_conf(self):
        self._app_enabled_list = pickle.load(open(self.enabled_apps_file, "rb"))
        self._update_app_disabled_list()

    def _write_apps_conf(self):
        pickle.dump(self._app_enabled_list, open(self.enabled_apps_file, "wb"))
        self._update_app_disabled_list()

    def addin_enable(self, addin_name):
        """
        Adds addin_name in the enabled addin list and save configuration.

        :param addin_name: string containing the name of the excel_addin to enable.
        :raise ConfigurationException: if add-in is already registered or unavailable is exceladdin sub-packages
        """
        if addin_name in self._addin_available_list:
            if addin_name in self._addin_enabled_list:
                mesg = "Configuration.addin_enable() : Addin {} is already enabled !".format(addin_name)
                print(mesg)
                raise ConfigurationException(mesg)
            else:
                print("Configuration.addin_enable() : Enabling {} addin...".format(addin_name))
                self._addin_enabled_list.append(addin_name)
                self._write_addin_conf()

                #fire on_addin_enable event
                for listener in self._event_listener_list:
                    listener.on_addin_enable(addin_name)
        else:
            mesg = "Configuration.addin_enable() : Can't enable unavailable {} addin !".format(addin_name)
            print(mesg)
            raise ConfigurationException(mesg)

    def addin_disable(self, addin_name):
        """
        Removes addin_name from the enabled addin list and save configuration.

        :param addin_name: string containing the name of the excel_addin to disable
        :raise ConfigurationException: if add-in is not in the available add-in list
        """
        if addin_name in self._addin_enabled_list:
            print("Disabling {} add-in...".format(addin_name))
            self._addin_enabled_list.remove(addin_name)
            self._write_addin_conf()

            #fire on_addin_disable event
            for listener in self._event_listener_list:
                listener.on_addin_disable(addin_name)

        else:
            mesg = "Configuration.addin_disable() : Can't disable unregistered {} addin !".format(addin_name)
            print(mesg)
            raise ConfigurationException(mesg)

    def _update_addin_disabled_list(self):
        # get available sub-packages of exceladdins
        for _, name, is_package in pkgutil.iter_modules(self._pkg_addin_path):
            if is_package and name not in self._addin_enabled_list:
                self._addin_disabled_list.append(name)

    def _update_app_disabled_list(self):
        # get available sub-packages of excelapps
        for _, name, is_package in pkgutil.iter_modules(self._pkg_app_path):
            if is_package and name not in self._app_enabled_list:
                self._app_disabled_list.append(name)

    def addin_get_desc(self, addin_name):
        """
        Return the description of the given add-in

        :param addin_name: string containing the name of the add-in
        :rtype str, -1 on error
        """
        try:
            # Dynamic import of the package to load comments
            module = inspect.importlib.import_module("exceladdins.{}".format(addin_name))
            # return top comments of the package
            return inspect.getcomments(module)

        except Exception as e:
            print(e)
            return -1

    def addin_get_disabled_list(self):
        """
        Returns disabled add-in list
        :return: disabled add-in list
        """
        return self._addin_disabled_list

    def addin_get_enabled_list(self):
        """
        Returns enabled add-in list
        :return: enabled add-in list
        """
        return self._addin_enabled_list

    def addin_get_list(self):
        """
        Return a list of tuple [(addin_name, bool_status)] of available add-ins
        :return: list
        """
        addin_list = list()
        # get available sub-packages of exceladdins
        for _, name, is_package in pkgutil.iter_modules(self._pkg_addin_path):
            if is_package and name in self._addin_enabled_list:
                addin_list.append((name, True))
            elif is_package:
                addin_list.append((name, False))

        return addin_list

    def addin_get_status(self, addin_name):
        """
        Returns status of given module
        :param addin_name:
        :return bool
        """
        status = False
        addin_list = self.addin_get_list()
        for name, status in addin_list:
            if name == addin_name:
                break

        return status

    def addin_print_list(self):
        """
        Displays names, state (loaded or not) and descriptions of the available
        addin.
        """
        print("Available exceladdins  :")
        addin_list = self.addin_get_list()

        for name, status in addin_list:
            if status:
                print("\n* {} [ENABLED] : \n".format(name))
            else:
                print("\n* {} [DISABLED] : \n".format(name))

            print(self.addin_get_desc(name))

    def app_enable(self, app_name):
        """
        Enables app_name and save configuration.
        :param app_name: name of the excel_app to enable.
        """
        if app_name in self._app_available_list:
            if app_name in self._app_enabled_list:
                mesg = "Configuration.app_enable() : app {} is already enabled !".format(app_name)
                print(mesg)
                raise ConfigurationException(mesg)
            else:
                print("Enabling {} app...".format(app_name))
                self._app_enabled_list.append(app_name)
                self._write_apps_conf()

                #fire on_addin_enable event
                for listener in self._event_listener_list:
                    listener.on_app_enable(app_name)
        else:
            mesg = "Configuration.app_enable() : Can't enable unavailable {} app !\nAvailable apps are {}.".format(
                app_name, repr(self._app_available_list))
            print(mesg)
            raise ConfigurationException(mesg)

    def app_disable(self, app_name):
        """
        Disables app_name and save configuration
        :param app_name: name of the excel_app to disable
        """
        if app_name in self._app_enabled_list:
            print("Disabling {} app...".format(app_name))
            self._app_enabled_list.remove(app_name)
            self._write_apps_conf()

            #fire on_addin_enable event
            for listener in self._event_listener_list:
                listener.on_app_disable(app_name)
        else:
            mesg = "Configuration.app_disable() : Can't disable unregistered {} app !".format(app_name)
            print(mesg)
            raise ConfigurationException(mesg)

    def app_get_desc(self, app_name):
        """
        Return the description of the given app

        :param app_name
        :rtype str
        :returns module description
        """

        try:
            # Dynamic import of the package - to be able to load comments
            module = inspect.importlib.import_module("excelapps.{}".format(app_name))
            # return top comments of the package
            return inspect.getcomments(module)

        except Exception as e:
            print(e)
            return -1

    def app_get_disabled_list(self):
        return self._app_disabled_list

    def app_get_enabled_list(self):
        return self._app_enabled_list

    def app_get_list(self):
        """
        Return a list of tuple [(addin_name, bool_status)] of available addins
        :rtype list
        """
        app_list = list()
        # get available sub-packages of excelapps
        for _, name, is_package in pkgutil.iter_modules(self._pkg_app_path):
            if is_package and name in self._app_enabled_list:
                app_list.append((name, True))
            elif is_package:
                app_list.append((name, False))

        return app_list

    def app_get_status(self, app_name):
        """
        Returns status of given module
        :param app_name
        :return boolean as status
        """
        status = False
        app_list = self.app_get_list()
        for name, status in app_list:
            if name == app_name:
                break

        return status

    def app_print_list(self):
        """
        Displays names, state (loaded or not) and descriptions of the available
        excelapps.

        """

        print("Available excelapps  :")

        app_list = self.app_get_list()

        for name, status in app_list:
            if status:
                print("\n* {} [ENABLED] : \n".format(name))
            else:
                print("\n* {} [DISABLED] : \n".format(name))

            print(self.app_get_desc(name))