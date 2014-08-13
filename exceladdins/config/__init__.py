# ------------------------------------------------------------------------------
# Name:        exceladdins.config
# Purpose:     Global configuration. Displays a button to open the conf window.
#
# In a standard installation, this Configuration box can be opened
# separately by clicking on "Configuration" in ExcelCOM folder
# in the Windows Start Menu.
# This file is responsible for Excel integration of the config box.
#
# Author:      Jonathan Besanceney <jonathan.besanceney@gmail.com>
#
# Created:     07/04/2014
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
In a standard installation, this Configuration box can be opened
separately by clicking on "Configuration" in Transition folder
in the Windows Start Menu.
This file is responsible for Excel integration of the config box.
"""
import sys

import os
from subprocess import Popen

sys.path += os.curdir
# specify python.exe to avoid drama creating new process with Excel
sys.executable = os.path.join(sys.exec_prefix, 'pythonw.exe')

#from PySide import QtGui  # Import PySide classes

from win32com.client import DispatchWithEvents, constants
import win32event

from exceladdins import addin
from exceladdins.config import config_box, configmain


class ButtonEvent:
    """ Button Event Handler class.
    This class is plugged to the button made in ConfigAddin Class in this file
    """

    def __init__(self):
        self.dialog = None
        self.script_name = "{}{}".format(os.path.abspath(os.path.dirname(__file__)), "\\configmain.py")

    def OnClick(self, button, cancel):
        #check if we already launched the dialog box
        if self.dialog is not None:
            #check if the dialog still running (Poll() return None)
            if self.dialog.poll() is not None:
                self.dialog = Popen([sys.executable, self.script_name])
            else:
                print("Dialog is already launched !")
        else:
            self.dialog = Popen([sys.executable, self.script_name])

        return cancel


class ConfigAddin(addin.ExcelAddin):
    def __init__(self, xl_app):
        super(ConfigAddin, self).__init__(xl_app, "ConfigAddin")
        self.cbcMyBar = self.xl_app.CommandBars.Add(Name="Transition Add-in conf",
                                              Position=constants.msoBarTop,
                                              MenuBar=constants.msoBarTypeNormal,
                                              Temporary=True)


        self.btnMyButton = self.cbcMyBar.Controls.Add(Type=constants.msoControlButton,
                                            Parameter="Greetings")

    def run(self):
        self.btnMyButton = DispatchWithEvents(self.btnMyButton, ButtonEvent)
        self.btnMyButton.Style = constants.msoButtonIconAndCaptionBelow
        self.btnMyButton.BeginGroup = True
        self.btnMyButton.Caption = "Transition config"
        self.btnMyButton.TooltipText = "Launch Transition config panel"
        self.btnMyButton.Width = "34"
        self.btnMyButton.FaceId = "642"
        self.btnMyButton.xlApp = self.xl_app

        self.cbcMyBar.Visible = True

        print(self.name, "running...")

    def wait(self):
        pass

    def resume(self):
        pass

    def terminate(self):
        #Does dialog still running ?
        if self.btnMyButton.dialog is not None:
            #check if the dialog still running (Poll() return None)
            if self.btnMyButton.dialog.poll() is None:
                self.btnMyButton.dialog.terminate()

        self.cbcMyBar.Visible = False
        self.cbcMyBar = None
        self.btnMyButton = None
        print(self.name, "terminated")

#declare our add-in. ExcelAddinManager will search module.excel_addin attribute to start this add-in
excel_addin = ConfigAddin


if __name__ == "__main__":
    script_name = "{}{}".format(os.path.abspath(os.path.dirname(__file__)), "\\configmain.py")
    print("try to launch {} {}".format(sys.executable, script_name))
    p = Popen([sys.executable, script_name])
    print(p.pid)
    p.wait()
