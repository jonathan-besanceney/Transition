# ------------------------------------------------------------------------------
# Name:        transitioncore
# Purpose:     Provides some constants to all others scripts.
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
 Provides :
 - some constants to all others scripts.

"""
from enum import Enum
import pythoncom

import exceladdins
import excelapps

defaultNamedNotOptArg = pythoncom.Empty
defaultMissingArg = pythoncom.Missing

class TransitionAppType(Enum):
    excel_addin = "exceladdins"  # python package name (used in dynamic imports)
    excel_wbapp = "excelapps"

transition_app_type_tree = {"excel" : {"addin" : TransitionAppType.excel_addin,
                                       "app" : TransitionAppType.excel_wbapp}}

transition_app_path = {TransitionAppType.excel_wbapp: excelapps.__path__ ,
                        TransitionAppType.excel_addin: exceladdins.__path__}