# ------------------------------------------------------------------------------
# Name:        addin_skell
# Purpose:      Define a standard way to replace VBA in Excel Add-in by an
#               external COM application launched by Transition Excel/COM Add-in.
#
# Author:      Jonathan Besanceney <jonathan.besanceney@gmail.com>
#
# Created:     09/04/2014
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
Define a standard way to replace VBA in Excel Add-in by an
external COM application launched by Transition Excel/COM Add-in.
"""
import sys

# specify free threading, common way to think threading.
sys.coinit_flags = 0

from threading import Thread


class ExcelAddinSkell(Thread):
    """Addin Standard Skeleton"""

    def __init__(self, xl_app, name="ExcelAddinSkell"):
        super(ExcelAddinSkell, self).__init__()
        self.name = name
        print("{} : Init".format(self.name))
        self.xl_app = xl_app
        self.ask_quit = False
        self.evt_handler = None

    def run(self):
        raise NotImplementedError("Subclasses must implement this method !")

    def quit(self):
        """
        Turns on the "quit" flag to stop event loop.
        :return:
        """
        self.ask_quit = True

if __name__ == '__main__':
    print("This file is a part of ExcelCOM project and is not intended to be run separately.")