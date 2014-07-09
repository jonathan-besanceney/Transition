# ------------------------------------------------------------------------------
# Name:        exceladdins
# Purpose:      exceladdins package
#
#               Provides a standard way to plug add-ins to Excel.
#               An add-in is a sub-program which always run with Excel and
#               extends its functionalities.
#
# Author:      Jonathan Besanceney <jonathan.besanceney@gmail.com>
#
# Created:     04/04/2014
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
Provides a standard way to plug add-ins to Excel.
An add-in is a sub-program which always run with Excel and
extends its functionalities.
"""
import inspect
import sys
import os

sys.path.append(os.path.abspath(os.path.dirname(__file__)))
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir)))

import transitionconfig


def get_desc():
    """Return the description of this module"""

    try:
        # Dynamic import of the package - to be able to load comments
        #inspect.importlib.import_module("excelapps")
        # return top comments of the package
        return inspect.getcomments(sys.modules["exceladdins"])

    except Exception as e:
        print(e)
        return -1


def unload_addins(addin_thread_list):

    """
    Stops all launched addins.
    :param addin_thread_list: list fo launched addins to stop
    """
    for addin in addin_thread_list:
        if addin.is_alive():
            addin.quit()

if __name__ == '__main__':
    print("This file is a part of ExcelCOM project and is not intended to be run separately.")


def load_addins(xlApp):
    """
    Loads and starts registered addins in the current xlApp

    :param xlApp: Excel.Application instance.

    :return the list of addin Threads

    """

    addin_list = transitionconfig.addin_get_enabled_list()

    print("exceladdins.load_addins() : List of addin to launch :", ' '.join(x for x in addin_list if x))

    addin_thread_list = []
    for name in addin_list:
        # Import app dynamicaly
        excel_addin_module = inspect.importlib.import_module("exceladdins.{}".format(name))
        addin = excel_addin_module.ExcelAddin(xlApp)
        addin.daemon = True
        addin.start()
        addin_thread_list.append(addin)

    return addin_thread_list