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

# specify free threading, common way to think threading.
sys.coinit_flags = 0

import os
from subprocess import Popen

sys.path += os.curdir
# specify python.exe to avoid drama creating new process with Excel
sys.executable = os.path.join(sys.exec_prefix, 'pythonw.exe')

#from PySide import QtGui  # Import PySide classes

from win32com.client import DispatchWithEvents, constants
import win32event

from exceladdins import addinskell
from exceladdins.config import config_box, configmain


class ButtonEvent:
    """ Button Event Handler class.
    This class is plugged to the button made in ExcelAddin Class in this file
    """

    def __init__(self):
        self.event = win32event.CreateEvent(None, 0, 0, None)
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

        win32event.SetEvent(self.event)
        return cancel


class ExcelAddin(addinskell.ExcelAddinSkell):
    def __init__(self, xlApp):
        super(ExcelAddin, self).__init__(xlApp)

    def run(self):
        self.name = "ConfigAddin"

        self.evt_handler = ButtonEvent

        cbcMyBar = self.xlApp.CommandBars.Add(Name="Transition Add-in conf",
                                              Position=constants.msoBarTop,
                                              MenuBar=constants
                                              .msoBarTypeNormal,
                                              Temporary=True)

        btnMyButton = cbcMyBar.Controls.Add(Type=constants.msoControlButton,
                                            Parameter="Greetings")
        btnMyButton = DispatchWithEvents(btnMyButton, self.evt_handler)
        btnMyButton.Style = constants.msoButtonIconAndCaptionBelow
        btnMyButton.BeginGroup = True
        btnMyButton.Caption = "&Transition config"
        btnMyButton.TooltipText = "Launch Transition config panel"
        btnMyButton.Width = "34"
        btnMyButton.FaceID = 1713
        btnMyButton.xlApp = self.xlApp

        cbcMyBar.Visible = True

        print(self.name, "addin plugged. Waiting for events...")

        # Main loop
        while self.ask_quit is False:
            win32event.WaitForSingleObject(btnMyButton.event, 1000)

        #Does dialog still running ?
        if btnMyButton.dialog is not None:
            #check if the dialog still running (Poll() return None)
            if btnMyButton.dialog.poll() is None:
                btnMyButton.dialog.terminate()

        self.xlApp = None


if __name__ == "__main__":
    script_name = "{}{}".format(os.path.abspath(os.path.dirname(__file__)), "\\configmain.py")
    print("try to launch {} {}".format(sys.executable, script_name))
    p = Popen([sys.executable, script_name])
    print(p.pid)
    p.wait()
