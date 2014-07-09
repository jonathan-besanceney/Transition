# ------------------------------------------------------------------------------
# Name:         post_install - Transition
# Purpose:      Tries to update environment without having to log-off or
#               reboot.
#
# Author:      Jonathan Besanceney <jonathan.besanceney@gmail.com>
#
# Created:     17/05/14
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

from win32api import SendMessage
import win32con


def main():
    """
    Asks for all windows to update their env variable.
    See :
    - kb104011
    - https://code.google.com/p/cefpython/source/browse/cefpython/var/envpath_broadcast.py
    """
    print("Sending WM_SETTINGCHANGE message to all opened windows...")
    SendMessage(win32con.HWND_BROADCAST, win32con.WM_SETTINGCHANGE, 0, "Environment")
    print("WM_SETTINGCHANGE sent.")


if __name__ == '__main__':
    main()