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
 - Excel Add-in register/unregister methods. (from <ekoome@yahoo.com> Eric Koome's
  /win32com/demo/excelAddin.py)
"""
import pythoncom

defaultNamedNotOptArg = pythoncom.Empty
defaultMissingArg = pythoncom.Missing

#How many milliseconds we wait for event in main loop
WAIT_FOR_EVENT_MSEC = 1000


def transition_register(klass):
    import winreg

    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                           "Software\\Microsoft\\Office\\Excel\\Addins")
    subkey = winreg.CreateKey(key, klass._reg_progid_)
    winreg.SetValueEx(subkey, "CommandLineSafe", 0, winreg.REG_DWORD, 0)
    winreg.SetValueEx(subkey, "LoadBehavior", 0, winreg.REG_DWORD, 3)
    winreg.SetValueEx(subkey, "Description", 0, winreg.REG_SZ,
                      klass.addin_name)
    winreg.SetValueEx(subkey, "FriendlyName", 0, winreg.REG_SZ,
                      klass.addin_description)


def transition_unregister(klass):
    import winreg

    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER,
                         "Software\\Microsoft\\Office\\Excel\\Addins\\" +
                         klass._reg_progid_)
    except WindowsError:
        pass