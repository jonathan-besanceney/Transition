# ------------------------------------------------------------------------------
# Name:        eventdispatcher.py
# Purpose:     Dispatch registered events. This is a quite basic implementation.
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


class TransitionEventDispatcher():
    def __init__(self, event_interface):
        self._event_interface = event_interface
        self._event_listener = list()

    def add_event_listener(self, listener):
        if isinstance(listener, self._event_interface):
            self._event_listener.append(listener)
        else:
            print("TransitionEventDispatcher.add_event_listener : expecting", self._event_interface.__class__.name,
                  "got", listener.__class__.name)

    def del_event_listener(self, listener):
        try:
            self._event_listener.remove(listener)
        except ValueError:
            print("TransitionEventDispatcher.del_event_listener : Can't remove unregistered listener", repr(listener))

    def _fire_event(self, event_method, event_args):
        """
        Fire an arbitrary registered event in the _com_event_listeners list.
        Looks for event method before calling it to prevent call useless empty method

        :param event_method: String containing event method name
        :param event_args:
        """
        for event_listener in self._event_listener:
            if event_method in event_listener.event_list:
                try:
                    if isinstance(event_args, tuple):
                        getattr(event_listener, event_method)(*event_args)
                    else:
                        getattr(event_listener, event_method)(event_args)
                except AttributeError as ae:
                    print("TransitionEventDispatcher._fire_event: except AttributeError calling", event_method, ae)
                # except TypeError as te:
                #    print("TransitionEventDispatcher._fire_event: except TypeError calling", event_method, te)