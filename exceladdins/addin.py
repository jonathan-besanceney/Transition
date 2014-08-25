# ------------------------------------------------------------------------------
# Name:        addin_skell
# Purpose:      Define a standard way to replace VBA in Excel Add-in by an
# external COM application launched by Transition Excel/COM Add-in.
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

class ExcelAddin():
    """Addin Standard Interface
        Addin Subpackages must implement this interface.
    """

    def __init__(self, xl_app, name="ExcelAddin"):
        self.name = name
        print("{} : Init".format(self.name))
        self.xl_app = xl_app

    def run(self):
        """Puts Excel Addin in run state"""
        raise NotImplementedError("Subclasses must implement this method !")

    def wait(self):
        """Puts Excel Addin in sleep state"""
        raise NotImplementedError("Subclasses must implement this method !")

    def resume(self):
        """Puts Excel Addin in run state after nap"""
        raise NotImplementedError("Subclasses must implement this method !")

    def terminate(self):
        """Puts Excel Addin in terminate mode"""
        raise NotImplementedError("Subclasses must implement this method !")


# class ExcelAddinManager():
#     #TODO : remove this class when AppManager is ok
#     def __init__(self, excel_app):
#         self.config = Configuration()
#         self.addin_list = self.config.get_enabled_app_list(TransitionAppType.excel_addin)
#         self.addin_instance_list = list()
#         self.excel_app = excel_app
#
#     def start_addins(self):
#         """
#         Loads and starts all registered addins in the current excel_app
#         """
#         print("ExcelAddinManager.start_addins() : List of addin to launch :", ' '.join(x for x in self.addin_list if x))
#
#         for name in self.addin_list:
#             self.start_addin(name)
#
#     def start_addin(self, name):
#         """Load and start an add-in by name.
#         :param name: sub-package name of exceladdin package
#         """
#         # Import app dynamicaly
#         excel_addin_module = inspect.importlib.import_module("exceladdins.{}".format(name))
#         if hasattr(excel_addin_module, "excel_addin") and issubclass(excel_addin_module.excel_addin, ExcelAddin):
#             addin = excel_addin_module.excel_addin(self.excel_app)
#             addin.run()
#             self.addin_instance_list.append(addin)
#         else:
#             print("ExcelAddinManager.start_addin() :", name, "is not a valid Transition Excel Addin.")
#
#     def terminate_addins(self):
#         """
#         Terminates all launched addins.
#         """
#         i = len(self.addin_instance_list)
#         while i != 0:
#             addin = self.addin_instance_list.pop()
#             addin.terminate()
#             i -= 1
#
#     def terminate_addin(self, name):
#         pass

if __name__ == '__main__':
    print("This file is a part of ExcelCOM project and is not intended to be run separately.")





