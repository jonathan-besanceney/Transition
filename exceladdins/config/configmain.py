# ------------------------------------------------------------------------------
# Name:         config_main - Transition
# Purpose:      Opens the Transition Configuration Dialog
#
# Author:      Jonathan Besanceney <jonathan.besanceney@gmail.com>
#
# Created:     01/06/14
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
Opens the Transition Configuration Dialog
"""

import sys
import os

import threading

sys.path.append(os.path.abspath(os.path.dirname(__file__)))
sys.path.append(os.path.abspath(os.path.join(os.path.join(os.path.dirname(__file__), os.pardir), os.pardir)))

# from threading import Thread

import PySide
from PySide.QtCore import Slot
import pywintypes
import win32trace

import transitionconfig
import exceladdins
import excelapps

from exceladdins.config import config_box

APPLICATIONS = "Applications"
ADDINS = "Add-ins"


class ControlConfigDialog(PySide.QtGui.QDialog):
    """
    Class controling the behaviour of the config_box.
    """

    def __init__(self, parent=None):
        super(ControlConfigDialog, self).__init__(parent)
        self.ui = config_box.Ui_Dialog()
        self.ui.setupUi(self)

        # Create a model to store the two lists (app & addin)
        # https://deptinfo-ensip.univ-poitiers.fr/ENS/pyside-docs/PySide/QtGui/QStandardItemModel.html
        treeModel = PySide.QtGui.QStandardItemModel()
        root = parentItem = treeModel.invisibleRootItem()

        itemAppRoot = PySide.QtGui.QStandardItem(APPLICATIONS)
        parentItem.appendRow(itemAppRoot)

        app_list = transitionconfig.app_get_list()

        for app_name, status in app_list:
            item = PySide.QtGui.QStandardItem(app_name)
            parentItem = itemAppRoot
            parentItem.appendRow(item)

        itemAddinRoot = PySide.QtGui.QStandardItem(ADDINS)
        parentItem = root
        parentItem.appendRow(itemAddinRoot)

        addin_list = transitionconfig.addin_get_list()
        for addin_name, status in addin_list:
            item = PySide.QtGui.QStandardItem(addin_name)
            parentItem = itemAddinRoot
            parentItem.appendRow(item)

        self.ui.treeViewAvailableComponents.setHeaderHidden(True)
        self.ui.treeViewAvailableComponents.setModel(treeModel)

        # timers must be inited in the main thread ?
        # timer init for displaying debug info
        self.timer = PySide.QtCore.QTimer(self)
        self.timer.setInterval(50)
        self.timer.timeout.connect(self.trace_writer)
        self.timer.start()

        # signal=>slot connections
        self.ui.treeViewAvailableComponents.clicked.connect(self.display_functionality_info)
        self.ui.buttonActivation.toggled.connect(self.set_button_activation_text)

    @Slot()
    def set_status(self, status):
        """
            Updates buttonActivation text upon the status.
            Register or unregister Addin/App

            :rtype : None
            :param status:
            """
        model = self.ui.treeViewAvailableComponents.selectedIndexes()

        for item in model:
            if status:
                if item.parent().data() == APPLICATIONS:
                    func = transitionconfig.app_enable
                else:
                    func = transitionconfig.addin_enable
            else:
                if item.parent().data() == APPLICATIONS:
                    func = transitionconfig.app_disable
                else:
                    func = transitionconfig.addin_disable
            func(item.data())
            print("set_status {} : {} {}".format(status, item.parent().data(), item.data()))
            self.set_button_activation_text(status)

    @Slot()
    def set_button_activation_text(self, status):
        if status:
            self.ui.buttonActivation.setText("Activé")
        else:
            self.ui.buttonActivation.setText("Désactivé")

    def trace_writer(self):
        """
        Writes traces in textDebugInfo
        """

        self.ui.textDebugInfo.insertPlainText(win32trace.blockingread(50))

    @Slot()
    def display_functionality_info(self, QModelIndex):
        self.ui.buttonActivation.toggled.disconnect(self.set_status)

        # we are on a root node
        if QModelIndex.parent().data() is None:
            self.ui.buttonActivation.setEnabled(False)
            self.ui.buttonActivation.setText("Statut")
            if QModelIndex.data() == APPLICATIONS:
                desc = excelapps.get_desc()
                self.ui.labelTitle.setText(APPLICATIONS)
            else:
                desc = exceladdins.get_desc()
                self.ui.labelTitle.setText(ADDINS)

        else:
            self.ui.buttonActivation.setEnabled(True)
            module_name = QModelIndex.data()
            if QModelIndex.parent().data() == APPLICATIONS:
                desc = transitionconfig.app_get_desc(module_name)
                status = transitionconfig.app_get_status(module_name)
            else:
                desc = transitionconfig.addin_get_desc(module_name)
                status = transitionconfig.addin_get_status(module_name)

            # toggle the button to reflect status and ensure status text is updated
            self.ui.buttonActivation.setChecked(status)
            self.set_button_activation_text(status)

            self.ui.labelTitle.setText(module_name)

        desc = desc.replace('#', '')
        desc = desc.replace("--", "")
        desc = desc.replace("-*- coding: utf8 -*-", "")
        self.ui.plainTextDescription.setPlainText(desc)

        self.ui.buttonActivation.toggled.connect(self.set_status)


def main():
    # qtapp = PySide.QtGui.QApplication.instance()
    #if qtapp is None:
    try:
        qtapp = PySide.QtGui.QApplication(sys.argv)
    except RuntimeError as e:
        print(e)

    qtapp.setStyle("plastique")

    dialog = ControlConfigDialog()

    try:
        win32trace.InitRead()
    except pywintypes.error as e:
        print(e)

    dialog.show()
    dialog.raise_()
    qtapp.exec_()

    try:
        win32trace.TermRead()
    except pywintypes.error as e:
        print(e)


if __name__ == "__main__":
    # main() => Qt: Could not initialize OLE (error 80010106) and python crashes after exec on win XP
    #Well here QT claims QApplication was not created in the main() thread, but python doesn't crash
    t = threading.Thread(target=main, name="ConfigDialog")
    t.start()
    t.join()