# ------------------------------------------------------------------------------
# Name:        transitionconfig
# Purpose:     transitionconfig package
#
#              Transition configuration manager.
#
# Author:      Jonathan Besanceney <jonathan.besanceney@gmail.com>
#
# Created:     21/06/2014
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
 Transition configuration manager.
 Provides :
 - addin/apps enable/disable methods
 - addin/apps list enable/disable methods
 - get descriptions and status of available addin/apps (based on inspect)
"""
import inspect
import pickle
import pkgutil
import sys
import os
from os.path import expanduser

sys.path.append(os.path.abspath(os.path.dirname(__file__)))
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir)))

import excelapps
import exceladdins

enabled_addins_file = expanduser("~") + "/.enabled_addins.list"
enabled_apps_file = expanduser("~") + "/.enabled_apps.list"

# Creates enabled_addins.list file if not exists
if not os.path.isfile(enabled_addins_file):
    new_addin_list = []
    pickle.dump(new_addin_list, open(enabled_addins_file, "wb"))

# Creates enabled_apps.list file if not exists
if not os.path.isfile(enabled_apps_file):
    handler_list = []
    pickle.dump(handler_list, open(enabled_apps_file, "wb"))


def addin_disable(addin_name):
    """
    Removes addin_name from the enabled addin list and save configuration.

    :param addin_name: string containing the name of the excel_addin to disable

    """

    addin_list = addin_get_enabled_list()

    if addin_name in addin_list:
        print("Disabling {} addin...".format(addin_name))
        addin_list.remove(addin_name)
        pickle.dump(addin_list, open(enabled_addins_file, "wb"))
    else:
        print("Cannot disable unregistered {} addin !".format(addin_name))


def addin_enable(addin_name):
    """
    Adds addin_name in the enabled addin list and save configuration.

    :param addin_name: string containing the name of the excel_addin to enable.

    """

    pkg_path = exceladdins.__path__

    app_exist = False
    # Check if addin exists
    for _, name, is_package in pkgutil.iter_modules(pkg_path):
        if is_package and addin_name == name:
            app_exist = True

    if app_exist:
        addin_list = addin_get_enabled_list()
        if addin_name in addin_list:
            print("Addin {} is already enabled !".format(addin_name))
        else:
            print("Enabling {} addin...".format(addin_name))
            addin_list.append(addin_name)
            pickle.dump(addin_list, open(enabled_addins_file, "wb"))
    else:
        print("Cannot enable unavailable {} addin !".format(addin_name))


def addin_get_desc(addin_name):
    """
    Return the description of the given add-in

    :param addin_name: string containing the name of the add-in
    :rtype str, -1 on error
    """

    try:
        # Dynamic import of the package - to be able to load comments
        module = inspect.importlib.import_module("exceladdins.{}".format(addin_name))
        # return top comments of the package
        return inspect.getcomments(module)

    except Exception as e:
        print(e)
        return -1


def addin_get_list():
    """
    Return a list of tuple [(addin_name, bool_status)] of available add-ins
    :rtype list
    """

    pkg_path = exceladdins.__path__

    addin_list = []
    enabled_addin_list = addin_get_enabled_list()

    # get available sub-packages of exceladdins
    for _, name, is_package in pkgutil.iter_modules(pkg_path):
        if is_package and name in enabled_addin_list:
            addin_list.append((name, True))
        elif is_package:
            addin_list.append((name, False))

    return addin_list


def addin_get_status(addin_name):
    """
    Returns status of given module
    :param addin_name:
    :rtype bool
    """
    status = False
    addin_list = addin_get_list()
    for name, status in addin_list:
        if name == addin_name:
            break

    return status


def addin_get_enabled_list():
    """
    Returns the list of enabled addins
    :rtype list
    """
    return pickle.load(open(enabled_addins_file, "rb"))


def addin_get_disabled_list():
    pkg_path = exceladdins.__path__

    addin_list = []
    enabled_addin_list = addin_get_enabled_list()

    # get available sub-packages of exceladdins
    for _, name, is_package in pkgutil.iter_modules(pkg_path):
        if is_package and name not in enabled_addin_list:
            addin_list.append(name)

    return addin_list


def addin_print_list():
    """
    Displays names, state (loaded or not) and descriptions of the available
    addin.
    :rtype None
    """
    print("Available exceladdins  :")
    addin_list = addin_get_list()

    for name, status in addin_list:
        if status:
            print("\n* {} [ENABLED] : \n".format(name))
        else:
            print("\n* {} [DISABLED] : \n".format(name))

        print(addin_get_desc(name))


def app_disable(app_name):
    """Disable app_name and save configuration

    Give it the name of the excel_app to disable

    :rtype None
    """

    app_enabled_list = app_get_enabled_list()

    if app_name in app_enabled_list:
        print("Disabling {} app...".format(app_name))
        app_enabled_list.remove(app_name)
        pickle.dump(app_enabled_list, open(enabled_apps_file, "wb"))
    else:
        print("Cannot disable unregistered {} app !".format(app_name))


def app_enable(app_name):
    """Adds an excel_app in the handler list and save it.

     Give it the name of the excel_app to enable.

    :rtype : None
    """

    pkg_path = excelapps.__path__

    app_exist = False
    # Check if application exists
    for _, name, is_package in pkgutil.iter_modules(pkg_path):
        if is_package and app_name == name:
            app_exist = True

    if app_exist:
        enabled_app_list = app_get_enabled_list()
        if app_name in enabled_app_list:
            print("Application {} is already enabled !".format(app_name))
        else:
            print("Enabling {} app...".format(app_name))
            enabled_app_list.append(app_name)
            pickle.dump(enabled_app_list, open(enabled_apps_file, "wb"))
    else:
        print("Cannot enable unavailable {} app !".format(app_name))


def app_get_desc(app_name):
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


def app_get_list():
    """
    Return a list of tuple [(addin_name, bool_status)] of available addins
    :rtype list
    """

    pkg_path = excelapps.__path__

    app_list = []
    enabled_app_list = app_get_enabled_list()

    # get available sub-packages of excelapps
    for _, name, is_package in pkgutil.iter_modules(pkg_path):
        if is_package and name in enabled_app_list:
            app_list.append((name, True))
        elif is_package:
            app_list.append((name, False))

    return app_list


def app_get_status(app_name):
    """
    Returns status of given module

    :param app_name
    :rtype bool
    :return boolean as status
    """

    status = False
    app_list = app_get_list()
    for name, status in app_list:
        if name == app_name:
            break

    return status


def app_get_enabled_list():
    return pickle.load(open(enabled_apps_file, "rb"))


def app_get_disabled_list():
    pkg_path = excelapps.__path__

    app_list = []
    enabled_app_list = app_get_enabled_list()

    # get available sub-packages of excelapps
    for _, name, is_package in pkgutil.iter_modules(pkg_path):
        if is_package and name not in enabled_app_list:
            app_list.append(name)

    return app_list


def app_print_list():
    """Displays names, state (loaded or not) and descriptions of the available
    excelapps.

    """

    print("Available excelapps  :")

    app_list = app_get_list()

    for name, status in app_list:
        if status:
            print("\n* {} [ENABLED] : \n".format(name))
        else:
            print("\n* {} [DISABLED] : \n".format(name))

        print(app_get_desc(name))





