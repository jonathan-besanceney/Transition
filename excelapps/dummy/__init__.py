# ------------------------------------------------------------------------------
# Name:        excelapps.dummy
# Purpose:     Dummy application for testing.
#
# Author:      Jonathan Besanceney <jonathan.besanceney@gmail.com>
#
# Created:     28/03/2014
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
Purpose:     Dummy application for testing.
"""


import sys

# specify free threading, common way to think threading.
sys.coinit_flags = 0

import os

sys.path.append(os.path.abspath(os.path.dirname(__file__)))
sys.path.append(os.path.abspath(os.path.join(os.path.join(os.path.dirname(__file__), os.pardir), os.pardir)))

from excelapps import appskell


class DummyApp(appskell.ExcelWorkbookAppSkell):
    """This class is a dummy workbook handler. It's main purpose is to be
    copied in a new excelapps sub package to make a real app.

    """

    def __init__(self, wb, evt_handler):
        super(DummyApp, self).__init__(wb, evt_handler,  "DummyApp")

    @staticmethod
    def is_handled_workbook(wb):
        """This method returns an instance of MyExcelWorkbookApp if wb is a
        handled Workbook.
        Implement here your workbook recognition.

        """
        return DummyApp(wb, appskell.ExcelWbEventsSkell)


if __name__ == '__main__':
    from win32com.client import Dispatch
    xlApp = Dispatch("Excel.Application")
    xlApp.Visible = 1
    m_wb = xlApp.Workbooks.Add()
    app = DummyApp.is_handled_workbook(m_wb)
    app.start()
    app = None
    xlApp = None
    sys.exit(0)