# ------------------------------------------------------------------------------
# Name:         documentapp
# Purpose:      documentapp package
#
# Provide a standard way to extend workbooks functionalities.
#
# The major benefit of it is to avoid code dispersion in many
# Excel workbooks in providing a unique code storage place.
#               There are also many amazing extensions capabilities like :
#                   - multi-thread your workbook code
#                   - plug your excel workbooks with your IT environment (DBs,
#                     web-services, outlook...)
#                   - extend interfacing with QT framework https://qt-project.org/
#                   - drive excel sheets easily with Microsoft Pyvot Python
#                   package. http://pytools.codeplex.com/wikipage?title=Pyvot
#                   - ...
#
#               Based on the high quality Mark Hammond's PyWin32 Python Package.
#               http://sourceforge.net/projects/pywin32/
#
# Author:       Jonathan Besanceney <jonathan.besanceney@gmail.com>
#
# Created:      28/03/2014
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
Provide a standard way to extend workbooks functionalities.

The major benefit of it is to avoid code dispersion in many
Excel workbooks in providing a unique code storage place.
              There are also many amazing extensions capabilities like :
                  - multi-thread your workbook code
                  - plug your excel workbooks with your IT environment (DBs,
                    web-services, outlook...)
                  - extend interfacing with QT framework https://qt-project.org/
                  - drive excel sheets easily with Microsoft Pyvot Python
                  package. http://pytools.codeplex.com/wikipage?title=Pyvot
                  - ...

              Based on the high quality Mark Hammond's PyWin32 Python Package.
              http://sourceforge.net/projects/pywin32/
"""

import inspect
import sys
import os

sys.path.append(os.path.abspath(os.path.dirname(__file__)))
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir)))

from documentapp.appskell import ExcelWorkbookAppSkell


def get_desc():
    """Return the description of this module"""

    try:
        # Dynamic import of the package - to be able to load comments
        #inspect.importlib.import_module("documentapp")
        # return top comments of the package
        return inspect.getcomments(sys.modules["documentapp"])

    except Exception as e:
        print(e)
        return -1


def get_wb_app_instance(wb):
    """Return the excel_app Thread instance to start().

    Give it a workbook instance to help this function to choose the good
    handler.

    Return None if no handler is enabled for the given workbook.

    """

    handler = None
    handler_list = app_get_enabled_list()
    for name in handler_list:
        # Import app dynamicaly
        excel_app_module = inspect.importlib.import_module("documentapp.{}".format(name))

        handler = excel_app_module.is_handled_workbook(wb)
        if handler is not None:
            continue

    return handler


