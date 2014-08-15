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


class ConfigEventsInterface():
    def on_app_enable(self, app_type, app_name):
        """
        Fired when an available app is enabled
        :param app_type: TransitionAppType Enum
        :param app_name:
        """
        pass

    def on_app_disable(self, app_type, app_name):
        """
        Fired when an available app is disabled
        :param app_type: TransitionAppType Enum
        :param app_name:
        """
        pass

    def on_app_add(self, app_type, app_name):
        """Fired when an app is added to available app list
        :param app_type: TransitionAppType Enum
        :param app_name:
        """
        pass

    def on_app_remove(self, app_type, app_name):
        """Fired when an app is removed from available app list
        :param app_type: TransitionAppType Enum
        :param app_name:
        """
        pass