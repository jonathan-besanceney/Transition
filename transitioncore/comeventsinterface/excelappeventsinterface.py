# ------------------------------------------------------------------------------
# Name:         excelappevents
# Purpose:      Interface used to define all events which can be handled by
#               Python/pywin32.win32com.client with :
#               * DispatchWithEvents("Excel.Application", ExcelAppEventsInterface)
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
  * DispatchWithEvents("Excel.Application", ExcelAppEventsInterface)
"""
import win32event

from transitioncore import defaultNamedNotOptArg, defaultMissingArg


class ExcelAppEventsInterface:
    """ExcelAppEventsInterface is an interface representing all Excel 2010 events.

    You can implement this class in your own Event Handler class or just take
    method you need. It's used like that :
        app = DispatchWithEvents("Excel.Application", ExcelAppEventsInterface)

    your app instance will be modified to include defined events
    in ExcelAppEventsInterface

    Note that Excel Event Handling via COM does not modify the behaviour of
    VBA in workbooks opened. You just need to be careful to the
    Application.EnableEvents property value.
    """

    def __init__(self):
        self.event = win32event.CreateEvent(None, 0, 0, None)
        self.name = 'ExcelAppEventsInterface'

    def OnAddRef(self):
        """TO BE DOCUMENTED"""
        print("{} ExcelAppEventsInterface OnAddRef".format(self.name))
        win32event.SetEvent(self.event)

    def OnAfterCalculate(self):
        print("{} ExcelAppEventsInterface OnAfterCalculate".format(self.name))
        win32event.SetEvent(self.event)

    def OnGetIDsOfNames(self, riid=defaultNamedNotOptArg, rgszNames=defaultNamedNotOptArg, cNames=defaultNamedNotOptArg,
                        lcid=defaultNamedNotOptArg, rgdispid=defaultNamedNotOptArg):
        print("{} ExcelAppEventsInterface OnGetIDsOfNames {}"
              .format(self.name, riid, rgszNames))
        win32event.SetEvent(self.event)

    def OnGetTypeInfo(self, itinfo=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, pptinfo=defaultMissingArg):
        print("{} ExcelAppEventsInterface OnGetTypeInfo {}"
              .format(self.name, repr(itinfo)))
        win32event.SetEvent(self.event)

    def OnGetTypeInfoCount(self, pctinfo=defaultMissingArg):
        print("{} ExcelAppEventsInterface OnGetTypeInfoCount {}"
              .format(self.name, repr(pctinfo)))
        win32event.SetEvent(self.event)

    def OnInvoke(self, dispidMember=defaultNamedNotOptArg, riid=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg,
                 wFlags=defaultNamedNotOptArg, pdispparams=defaultNamedNotOptArg, pvarResult=defaultMissingArg,
                 pexcepinfo=defaultMissingArg, puArgErr=defaultMissingArg):
        print("{} ExcelAppEventsInterface OnInvoke".format(self.name))
        win32event.SetEvent(self.event)

    def OnNewWorkbook(self, Wb=defaultNamedNotOptArg):
        """Occurs when a new workbook is created."""
        print("{} ExcelAppEventsInterface OnNewWorkbook {}".format(self.name, Wb.Name))
        win32event.SetEvent(self.event)

    def OnProtectedViewWindowActivate(self, Pvw=defaultNamedNotOptArg):
        print("{} ExcelAppEventsInterface OnProtectedViewWindowActivate".format(self.name))
        win32event.SetEvent(self.event)

    def OnProtectedViewWindowBeforeClose(self, Pvw=defaultNamedNotOptArg, Reason=defaultNamedNotOptArg,
                                         Cancel=defaultNamedNotOptArg):
        print("{} ExcelAppEventsInterface OnProtectedViewWindowBeforeClose {}"
              .format(self.name, Reason))
        win32event.SetEvent(self.event)
        return Cancel

    def OnProtectedViewWindowBeforeEdit(self, Pvw=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
        print("{} ExcelAppEventsInterface OnProtectedViewWindowBeforeEdit {}".format(self.name, Pvw))
        win32event.SetEvent(self.event)
        return Cancel

    def OnProtectedViewWindowDeactivate(self, Pvw=defaultNamedNotOptArg):
        print("{} ExcelAppEventsInterface OnProtectedViewWindowDeactivate {}".format(self.name, Pvw))
        win32event.SetEvent(self.event)

    def OnProtectedViewWindowOpen(self, Pvw=defaultNamedNotOptArg):
        print("{} ExcelAppEventsInterface OnProtectedViewWindowOpen {}".format(self.name, Pvw))
        win32event.SetEvent(self.event)

    def OnProtectedViewWindowResize(self, Pvw=defaultNamedNotOptArg):
        print("{} ExcelAppEventsInterface OnProtectedViewWindowResize {}"
              .format(self.name, Pvw))
        win32event.SetEvent(self.event)

    def OnQueryInterface(self, riid=defaultNamedNotOptArg, ppvObj=defaultMissingArg):
        print("{} ExcelAppEventsInterface OnQueryInterface {} {}"
              .format(self.name, riid, ppvObj))

    def OnRelease(self):
        print("{} ExcelAppEventsInterface OnRelease".format(self.name))
        win32event.SetEvent(self.event)

    def OnSheetActivate(self, Sh=defaultNamedNotOptArg):
        """Occurs when any sheet is activated."""

        print("{} ExcelAppEventsInterface OnSheetActivate {}"
              .format(self.name, Sh.Name))
        win32event.SetEvent(self.event)

    def OnSheetBeforeDoubleClick(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg,
                                 Cancel=defaultNamedNotOptArg):
        """Occurs when any worksheet is double-clicked, before the default
        double-click action."""

        print("{} ExcelAppEventsInterface OnSheetBeforeDoubleClick {} {}"
              .format(self.name, Sh.Name, Target.Address))
        win32event.SetEvent(self.event)
        return Cancel

    def OnSheetBeforeRightClick(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg,
                                Cancel=defaultNamedNotOptArg):
        """Occurs when any worksheet is right-clicked, before the default
        right-click action.
        Note that if you return TRUE, popup menus will be disabled.
        """

        print("{} ExcelAppEventsInterface OnSheetBeforeRightClick {} {}"
              .format(self.name, Sh.Name, Target.Address))
        win32event.SetEvent(self.event)
        return Cancel

    def OnSheetCalculate(self, Sh=defaultNamedNotOptArg):
        """Occurs after any worksheet is recalculated or after any changed
        data is plotted on a chart."""

        print("{} ExcelAppEventsInterface OnSheetCalculate {}"
              .format(self.name, Sh.Name))
        win32event.SetEvent(self.event)

    def OnSheetChange(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
        """Occurs when cells in any worksheet are changed by the user or by
        an external link."""

        print("{} ExcelAppEventsInterface OnSheetChange {} {}".format(self.name, Sh.Name, Target.Address))
        win32event.SetEvent(self.event)

    def OnSheetDeactivate(self, Sh=defaultNamedNotOptArg):
        """Occurs when any sheet is deactivated."""

        print("{} ExcelAppEventsInterface OnSheetDeactivate {}"
              .format(self.name, Sh.Name))
        win32event.SetEvent(self.event)

    def OnSheetFollowHyperlink(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
        """Occurs when you click any hyperlink in a workbook."""
        print("{} ExcelAppEventsInterface OnSheetFollowHyperlink {} {}"
              .format(self.name, Sh.Name, Target.Address))
        win32event.SetEvent(self.event)

    def OnSheetPivotTableAfterValueChange(self, Sh=defaultNamedNotOptArg, TargetPivotTable=defaultNamedNotOptArg,
                                          TargetRange=defaultNamedNotOptArg):
        """Occurs after a cell or range of cells inside a PivotTable are
        edited or recalculated (for cells that contain formulas).
        This event can be used only in Excel 2010 projects."""

        print("{} ExcelAppEventsInterface OnSheetPivotTableAfterValueChange {}"
              .format(self.name, Sh.Name))
        win32event.SetEvent(self.event)

    def OnSheetPivotTableBeforeAllocateChanges(self, Sh=defaultNamedNotOptArg, TargetPivotTable=defaultNamedNotOptArg,
                                               ValueChangeStart=defaultNamedNotOptArg,
                                               ValueChangeEnd=defaultNamedNotOptArg
                                               , Cancel=defaultNamedNotOptArg):
        """Occurs before changes are applied to a PivotTable.
        This event can be used only in Excel 2010 projects."""

        print("{} ExcelAppEventsInterface OnSheetPivotTableBeforeDiscardChanges {}"
              .format(self.name, Sh.Name, TargetPivotTable))
        win32event.SetEvent(self.event)
        return Cancel

    def OnSheetPivotTableBeforeCommitChanges(self, Sh=defaultNamedNotOptArg, TargetPivotTable=defaultNamedNotOptArg,
                                             ValueChangeStart=defaultNamedNotOptArg,
                                             ValueChangeEnd=defaultNamedNotOptArg
                                             , Cancel=defaultNamedNotOptArg):
        """Occurs before changes are committed against the OLAP data source
        for a PivotTable.
        This event can be used only in Excel 2010 projects."""

        print("{} ExcelAppEventsInterface OnSheetPivotTableBeforeCommitChanges {}"
              .format(self.name, Sh.Name))
        win32event.SetEvent(self.event)
        return Cancel

    def OnSheetPivotTableBeforeDiscardChanges(self, Sh=defaultNamedNotOptArg, TargetPivotTable=defaultNamedNotOptArg,
                                              ValueChangeStart=defaultNamedNotOptArg,
                                              ValueChangeEnd=defaultNamedNotOptArg):
        """Occurs before changes to a PivotTable are discarded.
        This event can be used only in Excel 2010 projects."""

        print("{} ExcelAppEventsInterface OnSheetPivotTableBeforeDiscardChanges {}"
              .format(self.name, Sh.Name, TargetPivotTable))
        win32event.SetEvent(self.event)

    def OnSheetPivotTableUpdate(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
        """Occurs after the sheet of a PivotTable report has been updated."""

        print("{} ExcelAppEventsInterface OnSheetPivotTableUpdate {}"
              .format(self.name, Sh.Name, Target.Address))
        win32event.SetEvent(self.event)

    def OnSheetSelectionChange(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
        """Occurs when the selection changes on any worksheet.
        Does not occur if the selection is on a chart sheet."""

        print("{} ExcelAppEventsInterface OnSheetSelectionChange {} {}"
              .format(self.name, Sh.Name, Target.Address))
        win32event.SetEvent(self.event)

    def OnWorkbookActivate(self, Wb=defaultNamedNotOptArg):
        """Occurs when the workbook is activated."""

        print("{} ExcelAppEventsInterface OnWorkbookActivate {}"
              .format(self.name, Wb.Name))
        win32event.SetEvent(self.event)

    def OnWorkbookAddinInstall(self, Wb=defaultNamedNotOptArg):
        """Occurs when the workbook is installed as an add-in."""

        print("{} ExcelAppEventsInterface OnWorkbookAddinInstall {}"
              .format(self.name, Wb.Name))
        win32event.SetEvent(self.event)

    def OnWorkbookAddinUninstall(self, Wb=defaultNamedNotOptArg):
        """Occurs when the workbook is uninstalled as an add-in."""

        print("{} ExcelAppEventsInterface OnWorkbookAddinUninstall {}"
              .format(self.name, Wb.Name))
        win32event.SetEvent(self.event)

    def OnWorkbookAfterSave(self, Wb=defaultNamedNotOptArg, Success=defaultNamedNotOptArg):
        """Occurs after the workbook is saved. This event can be used only
        in Excel 2010 projects."""

        print("{} ExcelAppEventsInterface OnWorkbookAfterSave {}"
              .format(self.name, Wb.Name))
        win32event.SetEvent(self.event)

    def OnWorkbookAfterXmlExport(self, Wb=defaultNamedNotOptArg, Map=defaultNamedNotOptArg, Url=defaultNamedNotOptArg,
                                 Result=defaultNamedNotOptArg):
        """Occurs after Microsoft Office Excel saves or exports data from the
        workbook to an XML data file."""

        print("{} ExcelAppEventsInterface OnWorkbookAfterXmlImport {}".format(self.name, Wb.Name))
        win32event.SetEvent(self.event)

    def OnWorkbookAfterXmlImport(self, Wb=defaultNamedNotOptArg, Map=defaultNamedNotOptArg,
                                 IsRefresh=defaultNamedNotOptArg, Result=defaultNamedNotOptArg):
        """Occurs after an existing XML data connection is refreshed or after
        new XML data is imported into the workbook."""

        print("{} ExcelAppEventsInterface OnWorkbookAfterXmlImport {}".format(self.name, Wb.Name))
        win32event.SetEvent(self.event)

    def OnWorkbookBeforeClose(self, Wb=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
        """Occurs before the workbook closes. If the workbook has been changed,
        this event occurs before the user is asked to save changes."""

        print("{} ExcelAppEventsInterface OnWorkbookBeforeClose {}"
              .format(self.name, Wb.Name))
        win32event.SetEvent(self.event)
        return Cancel

    def OnWorkbookBeforePrint(self, Wb=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
        """Occurs before the workbook (or anything in it) is printed."""

        print("{} ExcelAppEventsInterface OnWorkbookBeforePrint {}"
              .format(self.name, Wb.Name))
        win32event.SetEvent(self.event)
        return Cancel

    def OnWorkbookBeforeSave(self, Wb=defaultNamedNotOptArg, SaveAsUI=defaultNamedNotOptArg,
                             Cancel=defaultNamedNotOptArg):
        """Occurs before the workbook is saved.
        It should return FALSE (the default) to enable normal processing
        to continue.
        """
        print("{} ExcelAppEventsInterface OnWorkbookBeforeSave {}"
              .format(self.name, Wb.Name))
        win32event.SetEvent(self.event)
        return Cancel

    def OnWorkbookBeforeXmlExport(self, Wb=defaultNamedNotOptArg, Map=defaultNamedNotOptArg, Url=defaultNamedNotOptArg,
                                  Cancel=defaultNamedNotOptArg):
        """Occurs before Microsoft Office Excel saves or exports data from
        the workbook to an XML data file."""

        print("{} ExcelAppEventsInterface OnWorkbookBeforeXmlExport {}".format(self.name, Wb.Name))
        win32event.SetEvent(self.event)
        return Cancel

    def OnWorkbookBeforeXmlImport(self, Wb=defaultNamedNotOptArg, Map=defaultNamedNotOptArg, Url=defaultNamedNotOptArg,
                                  IsRefresh=defaultNamedNotOptArg
                                  , Cancel=defaultNamedNotOptArg):
        """Occurs before an existing XML data connection is refreshed or
        before new XML data is imported into the workbook."""

        print("{} ExcelAppEventsInterface OnWorkbookBeforeXmlImport {}"
              .format(self.name, Wb.Name))
        win32event.SetEvent(self.event)
        return Cancel

    def OnWorkbookDeactivate(self, Wb=defaultNamedNotOptArg):
        """Occurs when the workbook is deactivated."""

        print("{} ExcelAppEventsInterface OnWorkbookDeactivate {}"
              .format(self.name, Wb.Name))
        win32event.SetEvent(self.event)

    def OnWorkbookNewChart(self, Wb=defaultNamedNotOptArg, Ch=defaultNamedNotOptArg):
        """Occurs when a new chart is created in the workbook."""

        print("{} ExcelAppEventsInterface OnWorkbookNewChart {}"
              .format(self.name, Ch))
        win32event.SetEvent(self.event)

    def OnWorkbookNewSheet(self, Wb=defaultNamedNotOptArg, Sh=defaultNamedNotOptArg):
        """Occurs when a new sheet is created in the workbook."""

        print("{} ExcelAppEventsInterface OnWorkbookNewSheet {}"
              .format(self.name, Wb.Name, Sh.Name))
        win32event.SetEvent(self.event)

    def OnWorkbookOpen(self, Wb=defaultNamedNotOptArg):
        """Occurs when the workbook is opened."""

        print("{} ExcelAppEventsInterface OnWorkbookOpen {}".format(self.name, Wb.Name))
        win32event.SetEvent(self.event)

    def OnWorkbookSync(self, Wb=defaultNamedNotOptArg, SyncEventType=defaultNamedNotOptArg):
        """Occurs when the local copy of a worksheet that is part of a Document
        Workspace is synchronized with the copy on the server."""

        print("{} ExcelAppEventsInterface OnWorkbookSync {}"
              .format(self.name, Wb.Name))
        win32event.SetEvent(self.event)

    def OnWorkbookRowsetComplete(self, Wb=defaultNamedNotOptArg, Description=defaultNamedNotOptArg,
                                 Sheet=defaultNamedNotOptArg, Success=defaultNamedNotOptArg):
        """Occurs when the user navigates through the recordset or invokes the
        rowset action on an OLAP PivotTable."""

        print("{} ExcelAppEventsInterface OnWorkbookRowsetComplete {}"
              .format(self.name, Wb.Name))
        win32event.SetEvent(self.event)

    def OnWorkbookPivotTableCloseConnection(self, Wb=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
        """Occurs after a PivotTable report closes the connection to its
        data source."""

        print("{} ExcelAppEventsInterface OnWorkbookPivotTableCloseConnection {} {}".format(self.name, Wb.Name, Target.Address))
        win32event.SetEvent(self.event)

    def OnWorkbookPivotTableOpenConnection(self, Wb=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
        """Occurs after a PivotTable report opens the connection to its
        data source."""

        print("{} ExcelAppEventsInterface OnWorkbookPivotTableOpenConnection {} {}"
              .format(self.name, Wb.Name, Target.Address))
        win32event.SetEvent(self.event)

    def OnWindowActivate(self, Wb=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
        """Occurs when any workbook window is activated."""
        print("{} ExcelAppEventsInterface OnWindowActivate {} {}"
              .format(self.name, Wb.Name, Wn.Caption))
        win32event.SetEvent(self.event)

    def OnWindowDeactivate(self, Wb=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
        """Occurs when any workbook window is deactivated."""
        print("{} ExcelAppEventsInterface OnWindowDeactivate {} {}"
              .format(self.name, Wb.Name, Wn.Caption))
        win32event.SetEvent(self.event)

    def OnWindowResize(self, Wb=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
        """Occurs when any workbook window is resized."""
        print("{} ExcelAppEventsInterface OnWindowResize {} {}"
              .format(self.name, Wb.Name, Wn.Caption))
        win32event.SetEvent(self.event)


