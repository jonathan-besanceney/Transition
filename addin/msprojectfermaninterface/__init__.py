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
# Transition is free software: you can redistribute it and/or modify
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


from win32com.client import gencache
# InternetExplorer.Application
ie_module = gencache.EnsureModule('{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}', 0, 1, 1)
# IHTMLDocument2
gencache.EnsureModule('{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B}', 0, 4, 0)


class DocumentListener():
    def __init__(self):
        self.msp = None

    def set_msp(self, msp):
        self.msp = msp

    def OnLoad(self):
        print("onload")
        self.msp.login()

    def OnDownloadComplete(self):
        print("onload")
        self.msp.login()


class MSProjectFermanInterface():
    def __init__(self):
        self.ie = None
        self.doc = None

    def wait_page_loaded(self):
        while self.ie.ReadyState != ie_module.constants.READYSTATE_COMPLETE:
            pass

    def run(self):
        from win32com.client import Dispatch, DispatchWithEvents

        self.ie = Dispatch("InternetExplorer.Application")
        self.ie.Visible = True

        self.ie.Navigate2("http://powersales.power.alstom.com/service_enu/start.swe")
        self.wait_page_loaded()

        self.doc = self.ie.Document
        self.login()
        self.wait_page_loaded()
        print("logged in !")
        self.run_query({"Project ID": "A-C0B6AR"})

    def login(self):
        form = self.doc.forms.item(0)
        form.tags("input").item(0).value = 'jbesanceney'
        form.tags("input").item(1).value = 'jbesanceney1'
        form.submit()

    def go_fs_project(self):
        pass

    def run_query(self, criteria):
        querybt = self.doc.getElementById("s_2_1_7_0_mb")
        querybt.click()

    def terminate(self):
        self.ie.quit()

#declare our add-in. AppManager will search module.app_class attribute to start this
app_class = MSProjectFermanInterface
com_app = ('ms project', )
app_type = 'addin'

if __name__ == "__main__":
    msp = MSProjectFermanInterface()
    msp.run()