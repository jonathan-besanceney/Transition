# ------------------------------------------------------------------------------
# Name:         excelwbevents
# Purpose:      Interface used to define all events which can be handled by
#               Python/pywin32.win32com.client with :
#               * DispatchWithEvents("Excel.Workbook", ExcelWbEventsInterface)
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
"""
 Interface used to define all events which can be handled by
 Python/pywin32.win32com.client with :
 * DispatchWithEvents("Excel.Workbook", ExcelWbEventsInterface)
"""
import win32event

from transitioncore import defaultNamedNotOptArg


class ExcelWbEventsInterface:
    """ExcelWbEventsInterface est une interface qui représente tous les événements
    au niveau d'un classeur Excel.

    Cette classe est à ré-implémenter complètement ou en partie dans votre code
    source. Elle sera ensuite instanciée et utilisée à l'aide de :
        DispatchWithEvents("Excel.Workbook", ExcelWbEventsInterface)

    L'interception d'un événement par le code Python ne change pas le
    comportement de l'interception des événements VBA/Excel et peuvent donc
    être utilisés conjointement.

    """

    def __init__(self):
        self.event = win32event.CreateEvent(None, 0, 0, None)
        self.name = 'ExcelWbEventsInterface'

    def OnSheetActivate(self, sh):
        """Occurs when the worksheet is activated."""

        print("{} ExcelWbEventsInterface OnSheetActivate".format(self.name, sh.Name))
        win32event.SetEvent(self.event)

    def OnSheetBeforeDoubleClick(self, sh, Target, Cancel):
        """Occurs when the worksheet is double-clicked, before the default
        double-click action."""

        print("{} ExcelWbEventsInterface OnSheetBeforeDoubleClick"
              .format(self.name, sh.Name, Target.Address))
        win32event.SetEvent(self.event)
        return Cancel

    def OnSheetBeforeRightClick(self, sh, Target, Cancel):
        """Occurs when the worksheet is right-clicked, before the default
        right-click action."""

        print("{} ExcelWbEventsInterface OnSheetBeforeRightClick  {} {}"
              .format(self.name, sh.Name, Target.Address))
        win32event.SetEvent(self.event)
        return Cancel

    def OnSheetCalculate(self, sh):
        """Occurs after the worksheet is recalculated."""

        print("{} ExcelWbEventsInterface OnSheetCalculate {}".format(self.name, sh.Name))
        win32event.SetEvent(self.event)

    def OnSheetChange(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
        """Occurs when something changes in the Worksheet cells."""

        print("{} ExcelAppEventsInterface OnSheetChange {} {}".format(self.name, Sh.Name, Target.Address))
        win32event.SetEvent(self.event)

    def OnSheetDeactivate(self, sh):
        """Occurs when the worksheet loses focus."""

        print("{} ExcelWbEventsInterface OnSheetDeactivate {}"
              .format(self.name, sh.Name))
        win32event.SetEvent(self.event)

    def OnSheetFollowHyperlink(self, sh, Target):
        """Occurs when you click any hyperlink on a worksheet."""

        print("{} ExcelWbEventsInterface OnSheetFollowHyperlink {} {}"
              .format(self.name, sh.Name, Target.Address))
        win32event.SetEvent(self.event)

    def OnSheetSelectionChange(self, sh, Target):
        """Occurs when the selection changes on a worksheet."""

        print("{} ExcelWbEventsInterface OnSheetSelectionChange {} {}"
              .format(self.name, sh.Name, Target.Address))
        win32event.SetEvent(self.event)

    def OnDeactivate(self):
        print("{} ExcelWbEventsInterface OnDeactivate".format(self.name))
        win32event.SetEvent(self.event)

    def OnBeforeClose(self, Cancel):
        print("{} ExcelWbEventsInterface OnBeforeClose".format(self.name))
        win32event.SetEvent(self.event)

    def OnWindowActivate(self, Wn):
        print("{} ExcelWbEventsInterface OnWindowActivate {}".format(self.name, Wn.Caption))
        win32event.SetEvent(self.event)

    def OnWindowResize(self, Wn):
        print("{} ExcelWbEventsInterface OnWindowResize {}".format(self.name, Wn.Caption))
        win32event.SetEvent(self.event)

    def OnRelease(self):
        print("{} ExcelWbEventsInterface OnRelease".format(self.name))
        win32event.SetEvent(self.event)

    def OnNewsheet(self, sh):
        print("{} ExcelWbEventsInterface OnNewsheet {}".format(self.name, sh.Name))
        win32event.SetEvent(self.event)

    def OnAddinUninstall(self):
        print("{} ExcelWbEventsInterface OnAddinUninstall".format(self.name))
        win32event.SetEvent(self.event)

    def OnAddinInstall(self):
        print("{} ExcelWbEventsInterface OnAddinInstall".format(self.name))
        win32event.SetEvent(self.event)

    def OnOpen(self):
        print("{} ExcelWbEventsInterface OnOpen".format(self.name))
        win32event.SetEvent(self.event)

    def OnAddRef(self):
        print("{} ExcelWbEventsInterface OnAddRef".format(self.name))
        win32event.SetEvent(self.event)

    def OnWindowDeactivate(self, Wn):
        print("{} ExcelWbEventsInterface OnWindowDeactivate".format(self.name, Wn.Caption))
        win32event.SetEvent(self.event)

    def OnActivate(self):
        print("{} ExcelWbEventsInterface OnActivate".format(self.name))
        win32event.SetEvent(self.event)

    def OnBeforePrint(self, Cancel):
        print("{} ExcelWbEventsInterface OnBeforePrint".format(self.name))
        win32event.SetEvent(self.event)

    def OnBeforeSave(self, SaveAsUI, Cancel):
        print("{} ExcelWbEventsInterface OnBeforeSave".format(self.name))
        win32event.SetEvent(self.event)