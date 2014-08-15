# ------------------------------------------------------------------------------
# Name:        Script Name 
# Purpose:     TODO 
#
# Author:      Jonathan Besanceney <jonathan.besanceney@gmail.com>
#
# Created:     01/06/14
# Copyright:   (c) 2014 Jonathan Besanceney
#
# This file is a part of Transition
#
#    Transition is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
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

from transitioncore.comeventsinterface.comeventsinterface import COMEventsInterface


class KernelComEventListener(COMEventsInterface):
    def __init__(self, kernel_instance=None):
        super(KernelComEventListener, self).__init__()
        self.event_list = ("on_connection", "on_begin_shutdown", "on_startup_complete")
        self.kernel_instance = kernel_instance

    def on_connection(self, application, connect_mode, addin, custom):
        # give application and add-in references to kernel
        self.kernel_instance.set_application(application)
        self.kernel_instance.set_addin(addin)

    def on_begin_shutdown(self, custom):
        self.kernel_instance.terminate()

    def on_startup_complete(self, custom):
        # Excel is up. Start Kernel
        self.kernel_instance.run()