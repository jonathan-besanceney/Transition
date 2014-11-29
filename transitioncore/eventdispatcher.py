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
    """
    Small event dispatcher.
    Subclass it to manage events in your own class.
    """

    def __init__(self, event_interface):
        # Declare event interface type. Used to control event listener registration.
        self._event_interface = event_interface
        # List of event listener
        self._event_listener = list()

    def add_event_listener(self, listener):
        """
        Add (register) an Event Listener class instance
        :param listener: event_interface subclass
        :return: None
        """
        print("given listener", repr(listener.__class__))
        print("expected listener", repr(self._event_interface))
        # if issubclass(type(listener), self._event_interface):
        self._event_listener.append(listener)
        # else:
        #     mesg = "TransitionEventDispatcher.add_event_listener : unexpected type"
        #     print(mesg)
        #     raise TypeError(mesg)
        print(len(self._event_listener), "listener(s) registered")

    def del_event_listener(self, listener):
        """
        Delete (unregister) previously registered Event Listener class instance.
        :param listener:
        :return:
        """
        try:
            self._event_listener.remove(listener)
            print(len(self._event_listener), "listener(s) registered")
        except ValueError:
            print("TransitionEventDispatcher.del_event_listener : Can't remove unregistered listener", repr(listener))

    def _fire_event(self, event_method, event_args):
        """
        Fire an arbitrary registered event in the _com_event_listeners list.
        Looks for event method before calling it to prevent call useless empty method

        :param event_method: String containing event method name
        :param event_args:
        """
        print("_fire_event called.")
        for event_listener in self._event_listener:
            print("_fire_event : look for ", event_method, "in", event_listener)
            if event_method in event_listener.event_list:
                print("_fire_event : fire", event_method, "with", event_args)
                try:
                    if isinstance(event_args, tuple):
                        getattr(event_listener, event_method)(*event_args)
                    else:
                        getattr(event_listener, event_method)(event_args)
                except AttributeError as ae:
                    print("TransitionEventDispatcher._fire_event: raise AttributeError calling", event_method, ae)
                # except TypeError as te:
                #    print("TransitionEventDispatcher._fire_event: except TypeError calling", event_method, te)
                print("_fire_event : fired :", event_method, "with", event_args)