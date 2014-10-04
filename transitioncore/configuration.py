# ------------------------------------------------------------------------------
# Name:        configuration
# Purpose:     Configuration Class
#
# Author:      Jonathan Besanceney <jonathan.besanceney@gmail.com>
#
# Created:     01/06/14
# Copyright:   (c) 2014 Jonathan Besanceney
#
# This file is a part of Transition
#
# Transition is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# Transition is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU Lesser General Public License for more details.
#
#    You should have received a copy of the GNU Lesser General Public License
#    along with Transition.  If not, see <http://www.gnu.org/licenses/>.
# ------------------------------------------------------------------------------
# -*- coding: utf8 -*-

import inspect
import os

import pkgutil
import sys
import zipfile
import hashlib
import sha3
import whirlpool
import pickle
import sqlite3

sys.path.append(os.path.abspath(os.path.dirname(__file__)))

from transitioncore.eventdispatcher import TransitionEventDispatcher
from transitioncore.eventsinterface.configeventinterface import ConfigEventsInterface
from transitioncore.exceptions.configurationexception import ConfigurationException
from transitioncore.configsql import *


class Configuration(TransitionEventDispatcher):
    """ Deals with Transition configuration.

    Application management
        * generate app digests
        * verify app digests against Manifest
        * verify app digests against db
        * read app manifest
        * generate app manifest
        * update app info on db
        * add app info on db
        * del app info from db
        * update app inventory and info (trigger update, add, del)

    Application Configuration :
        * enable app
        * disable app
        * reset configuration

    Application information :
        * get app list (from files / db)
        * get app description (from file / db)
        * get app status (db) enabled/disabled
        * get app mode (Manifest file) devmode/usermode
        * get app state (from digests) unchanged/updated/corrupted
        * ??? get available app list (from ?)
        * get disabled app (from db)
        * get enabled app (from db)

    Helpers :
        * get app id from its name and type
        * get com app id from its name
        * get app type path
        * get info from app desc
        * print app list to std out

    Notes on Digests (checksums) verification :
        * User mode (with Manifest file)
        generated app digests MUST be the same as Manifest. If not, application is corrupted (state) and is disabled
        (status) until Manifest is updated.
        when db digests differs from generated app digests, configuration updates db digests and fire on_app_update

        * Dev mode (without Manifest file)
        When db digests differs from generated app digests, configuration updates db digests and fire on_app_update
        event

        Notes about hash :
        * The Whirlpool hashing algorithm (http://www.larc.usp.br/~pbarreto/WhirlpoolPage.html),
        written by Vincent Rijmen and Paulo S. L. M. Barreto is a secure, modern hash which is as yet unbroken and
        fairly obscure.
        * SHA-3: A hash function formerly called Keccak, chosen in 2012 after a public competition
        among non-NSA designers. It supports the same hash lengths as SHA-2, and its internal structure
        differs significantly from the rest of the SHA family.

    Events
        * on_app_add
        * on_app_del
        * on_app_enable
        * on_app_disable
        * on_app_update
    """

    cnx_str = os.getenv("USERPROFILE") + "\\.transition.s3db"

    create_tables = (SQL_CREATE_APP, SQL_CREATE_APP_TYPE, SQL_CREATE_COM_APP, SQL_CREATE_APP_WORKS_WITH_COM_APP)

    app_type_list = ('docapp', 'addin')
    com_app_list = ('Excel', 'Access', 'MS Project', 'OneNote', 'Outlook', 'PowerPoint', 'Word')

    def _create_tables(self):
        """
        Creates configuration tables and populates them
        """
        cursor = self._sqlite_cnx.cursor()
        sql = ""
        try:
            for sql in Configuration.create_tables:
                print(sql)
                cursor.execute(sql)

            for com_app in Configuration.com_app_list:
                print(SQL_INSERT_COM_APP)
                cursor.execute(SQL_INSERT_COM_APP, (com_app.lower(), ))

            for app_type in Configuration.app_type_list:
                module = inspect.importlib.import_module(app_type)
                print(SQL_INSERT_APP_TYPE)
                cursor.execute(SQL_INSERT_APP_TYPE, (app_type, module.__path__[0]))

            self._sqlite_cnx.commit()
        except sqlite3.OperationalError as oe:
            print("Configuration._create_tables :", sql, oe)

    def reset(self):
        if self._sqlite_cnx is not None:
            self._sqlite_cnx.close()

        if os.path.exists(self.cnx_str):
            os.remove(self.cnx_str)

    def __init__(self, com_app=None):
        """
        Configuration
        :param com_app: if set, tells config we are linked to this com_app.
        :return:
        """
        super(Configuration, self).__init__(ConfigEventsInterface)

        self.com_app = com_app
        self._sqlite_cnx = None

        #events listeners
        self._event_listener_list = list()

    def _get_sqlite(self):
        """
        Make SQLite connection. Creates DB if not exists.
        :return: DB2API sqlite connection
        """

        db_create = False
        if not os.path.exists(Configuration.cnx_str):
            self._sqlite_cnx = None
            db_create = True

        if self._sqlite_cnx is None:
            self._sqlite_cnx = sqlite3.connect(Configuration.cnx_str)
            self._sqlite_cnx.row_factory = sqlite3.Row
            if db_create:
                self._create_tables()

            self.update_inventory(fire_event=False)

        return self._sqlite_cnx

    _sqlite = property(_get_sqlite)

    def get_app_type_path(self, app_type):
        """
        Return app type path from db
        :param app_type:
        :return: app type path, None if app_type doesn't exist
        """
        cursor = self._sqlite.cursor()
        cursor.execute(SQL_SELECT_APP_TYPE_PATH, (app_type, ))
        row = cursor.fetchone()
        if row is None:
            path = None
        else:
            path = row["PATH"]
        #print("get_app_type_path()", app_type, path)
        return path

    @staticmethod
    def _generate_app_digests(path):
        """
        Generate application digests in sha256, sha512 and whirlpool.
        Used to verify apps source integrity. Pycache is regenerated by AppManager.
        Idea comes from : http://devmanual.gentoo.org/general-concepts/manifest/index.html
        SHA uses keccak. See top comments about that.

        :param app_type: application type (package)
        :param app_name : application name (sub-package)
        :return: dict {SHA256, SHA512, WHIRLPOOL} or None on error
        """

        digests = None
        try:
            if os.path.exists(path):
                try:
                    source_path = os.getenv("TMP") + "\\transition_app_source.zip"
                    source_count = 0
                    zipf_source = zipfile.ZipFile(source_path, 'w')


                    #TODO : find a way to deal with pycache
                    for root, dirs, files in os.walk(path):
                        if "__pycache__" not in root:
                            for file in files:
                                if file not in ("Manifest",):
                                    source_count += 1
                                    zipf_source.write(os.path.join(root, file))

                    zipf_source.close()

                    if source_count > 0:
                        source_to_check = open(source_path, 'rb')
                        source_content = source_to_check.read()

                        source_sha256 = hashlib.sha3_256()
                        source_sha256.update(source_content)

                        source_sha512 = hashlib.sha3_512()
                        source_sha512.update(source_content)

                        source_wp = whirlpool.new()
                        source_wp.update(source_content)

                        source_to_check.close()

                        digests = {
                            'SHA256': source_sha256.hexdigest(),
                            'SHA512': source_sha512.hexdigest(),
                            'WHIRLPOOL': source_wp.hexdigest()
                        }
                    else:
                        print("_generate_app_digests() : no files found in path", path)

                    #print(source_path)
                    os.remove(source_path)
                except Exception as e:
                    print(repr(e))
            else:
                print("_generate_app_digests() : invalid path", path)
        except ImportError as ie:
            print("_generate_app_digests() : Import Error", repr(ie))

        return digests

    @staticmethod
    def write_app_manifest(path):
        """
        write a new Manifest file containing digests for app.
        :param path:
        :return: digests or None on error
        :raise ConfigurationException if Manifest file can't be created
        """

        try:
            digests = Configuration._generate_app_digests(path)
            if digests is not None:
                with open(path + '\\Manifest', 'wb') as f:
                    pickle.dump(digests, f)

            return digests
        except IOError:
            raise ConfigurationException("Can't open or create Manifest file here : {} {}".format(path, repr(IOError)))

    @staticmethod
    def _read_app_manifest(path):
        """
        Read Manifest file in path
        :param path: application path
        :return: digests dict
        :raise: ConfigurationException on error
        """
        manifest_file = path + '\\Manifest'
        ret_val = None
        if os.path.exists(manifest_file):
            try:
                with open(manifest_file, 'rb') as f:
                    digests = pickle.load(f)

                if isinstance(digests, dict):
                    keys = digests.keys()
                    if 'SHA256' in keys \
                            and 'SHA512' in keys \
                            and 'WHIRLPOOL' in keys:
                        if len(digests['SHA256']) == 64 \
                                and len(digests['SHA512']) == 128 \
                                and len(digests['WHIRLPOOL']) == 128:
                            ret_val = digests

                #if loaded object is not a valid dict containing SHA256, SHA512, WHIRLPOOL keys of
                #64, 128 and 128 bytes, raise exception
                if ret_val is None:
                    raise ConfigurationException("Invalid Manifest file in {}.".format(path))
                else:
                    return ret_val
            except EOFError as eofe:
                raise ConfigurationException("Manifest file seems empty {} : {}.".format(path, repr(eofe)))
        else:
            raise ConfigurationException("{} doesn't contain Manifest file.".format(path))

    def _verify_db_digests(self, path):
        """
        Compare digests stored in db file and generated digests
        :param path: app path (must contain Manifest File)
        :return: True/False , generated app digest/None
        """
        ret_val = True
        digests = Configuration._generate_app_digests(path)
        if digests is not None:
            db_digests = self.get_app_info_by_path(path)
            if db_digests is not None:
                for key in digests.keys():
                    if digests[key] != db_digests[key]:
                        ret_val = False
                        print("Configuration._verify_db_digests() : Stored", key, "in Manifest file", db_digests[key],
                              "doesn't match with actual", key, "signature", digests[key])
                    # else:
                    #     print("Configuration._verify_db_digests() :", key, "match", digests[key])
            else:
                ret_val = False
        else:
            ret_val = False

        return ret_val, digests

    @staticmethod
    def _verify_manifest_digests(path):
        """
        Compare digests stored in Manifest file and generated digests
        :param path: app path (must contain Manifest File)
        :return: true/false, generated digests / None if digests can't be generated
        """
        ret_val = True
        digests = Configuration._generate_app_digests(path)
        if digests is not None:
            try:
                manifest = Configuration._read_app_manifest(path)

                for key in digests.keys():
                    if digests[key] != manifest[key]:
                        ret_val = False
                        print("Configuration._verify_manifest_digests() : Stored", key, "in Manifest file",
                              manifest[key],
                              "doesn't match with actual", key, "signature", digests[key])
                    # else:
                    #     print("Configuration._verify_manifest_digests() :", key, "matches", digests[key])
            except ConfigurationException as ce:
                ret_val = False
                print("Configuration._verify_manifest_digests() : Can't verify signatures : ", ce.value)
        else:
            ret_val = False

        return ret_val, digests

    def get_app_info_by_path(self, path):
        """
        Return app db record.
        :param path: application path
        :return: row containing app fields, None if app doesn't exist
        """
        cursor = self._sqlite.cursor()
        cursor.execute(SQL_SELECT_APP_BY_PATH, (path, ))
        return cursor.fetchone()

    def get_app_info(self, app_type, app_name):
        """
        Return app db record.
        :param app_type: application type
        :param app_name: application name
        :return: row containing app fields, None if app doesn't exist
        """
        cursor = self._sqlite.cursor()
        cursor.execute(SQL_SELECT_APP, (app_type, app_name))
        return cursor.fetchone()

    def update_inventory(self, app_type=None, fire_event=True):
        """
        Updates available app list. Fire on_app_add(self, addin_name) or on_app_del(self, addin_name)
        on change
        :param app_type : update inventory for particular app_type (all types per default)
        :param fire_event : fire events per default. Set it to false to disable it.
        """
        if app_type is None:
            for app_type in self.app_type_list:
                self.update_inventory(app_type, fire_event)
        else:
            #print("update_inventory()", app_type)
            # get available app_type sub-packages watching for application
            app_path = self.get_app_type_path(app_type)
            app_list = list()
            for _, app_name, is_package in pkgutil.iter_modules((app_path, )):
                app_list.append(app_name)
                if is_package and app_name not in self.get_available_app_list(app_type):
                    self._app_add(app_type, app_name, fire_event)
                elif is_package:
                    # known app. Verify state.
                    state = self.get_app_state(app_type, app_name)
                    if state == -1:
                        self.disable_app(app_type, app_name)
                    elif state == 1:
                        digests = self._generate_app_digests(app_path + "\\" + app_name)
                        self._app_update(app_type, app_name, digests, fire_event)

            # now we want to see app deletion
            app_available_list = self.get_available_app_list(app_type)
            for name in app_available_list:
                if name not in app_list:
                    self._app_del(app_type, name, fire_event)

    def get_com_app_list(self):
        """Return com_app list from config
        :return: list() with com_app_name
        """
        cursor = self._sqlite.cursor()
        cursor.execute(SQL_SELECT_ALL_COM_APP)
        com_apps_name = list()
        for row in cursor:
            com_apps_name.append(row['short_name'])

        return com_apps_name

    def enable_app(self, app_type, app_name, com_app=None):
        """
        Enables app_name for specified com_app and save configuration.
        :param app_type: application type
        :param app_name: name of the excel_app to enable.
        :param com_app: com_app in witch we want to activate app.
        None per default to enable app for all com_app.
        :return: True if enabling at least on app/com_app succeed, False instead
        :raise: ConfigurationException if app_type or app doesn't exists
        """

        if app_type not in self.app_type_list:
            mesg = "Configuration.enable_app() : Given app_type {} doesn't exist !\nAvailable app_type are {}.".format(
                app_type, self.app_type_list)
            print(mesg)
            raise ConfigurationException(mesg)
        elif app_name not in self.get_available_app_list(app_type):
            mesg = "Configuration.enable_app() : Can't enable unavailable {} app !\nAvailable apps are {}.".format(
                app_name, repr(self.get_available_app_list(app_type)))
            print(mesg)
            raise ConfigurationException(mesg)
        else:
            #get app/com_app status
            app_status_list = self.get_app_list(app_type=app_type, app_name=app_name, com_app=com_app)

            #iterate over app_status_list
            enabled_com_app_tuple = list()
            for item in app_status_list:
                #skip enabled app/com_app
                if item['enabled'] == 1:
                    mesg = "Configuration.enable_app() : app {} is already enabled for {} !".format(app_name,
                                                                                                    com_app)
                    print(mesg)
                else:
                    print("Enabling {} app for {}...".format(app_name, item['com_app']))
                    cursor = self._sqlite.cursor()
                    cursor.execute(SQL_UPDATE_APP_WORKS_WITH_COM_APP, (True, app_type, app_name, item['com_app']))
                    self._sqlite.commit()
                    enabled_com_app_tuple.append(item['com_app'])

            if len(enabled_com_app_tuple) > 0:
                #fire on_addin_enable event
                self._fire_event("on_app_enable", (app_type, app_name, tuple(enabled_com_app_tuple)))
                ret_val = True
            else:
                ret_val = False

        return ret_val

    def disable_app(self, app_type, app_name, com_app=None):
        """
        Disable app_name for specified com_app and save configuration.
        :param app_type: application type
        :param app_name: name of the excel_app to disable.
        :param com_app: com_app in witch we want to activate app.
        None per default to disable app for all com_app.
        :return: True if disabling at least on app/com_app succeed, False instead
        :raise: ConfigurationException if app_type/app doesn't exists
        """
        if app_type not in self.app_type_list:
            mesg = "Configuration.enable_app() : Given app_type {} doesn't exist !\nAvailable app_type are {}.".format(
                app_type, self.app_type_list)
            print(mesg)
            raise ConfigurationException(mesg)
        elif app_name not in self.get_available_app_list(app_type):
            mesg = "Configuration.enable_app() : Can't enable unavailable {} app !\nAvailable apps are {}.".format(
                app_name, repr(self.get_available_app_list(app_type)))
            print(mesg)
            raise ConfigurationException(mesg)
        else:
            #get app/com_app status
            app_status_list = self.get_app_list(app_type=app_type, app_name=app_name, com_app=com_app)

            #iterate over app_status_list
            disabled_com_app_tuple = list()
            for item in app_status_list:
                #skip enabled app/com_app
                if item['enabled'] == 0:
                    mesg = "Configuration.disable_app() : app {} is already disabled for {} !".format(app_name,
                                                                                                    com_app)
                    print(mesg)
                else:
                    print("Disabling {} app for {}...".format(app_name, item['com_app']))
                    cursor = self._sqlite.cursor()
                    cursor.execute(SQL_UPDATE_APP_WORKS_WITH_COM_APP, (False, app_type, app_name, item['com_app']))
                    self._sqlite.commit()
                    disabled_com_app_tuple.append(item['com_app'])

            if len(disabled_com_app_tuple) > 0:
                #fire on_addin_enable event
                self._fire_event("on_app_disable", (app_type, app_name, tuple(disabled_com_app_tuple)))
                ret_val = True
            else:
                ret_val = False

        return ret_val

    def get_disabled_app_list(self, app_type, com_app):
        """
        Returns disabled app list for given app_type and com_app
        :param app_type:
        :param com_app : Specify com_app.
        :return: disabled app list, None if app_type and/or com_app doesn't exist
        """
        disabled_app_list = None
        if app_type in self.app_type_list and com_app in self.get_com_app_list():
            app_list = self.get_app_list(app_type=app_type, com_app=com_app, enabled=False)
            if app_list is None:
                app_list = list()

            disabled_app_list = list()
            for row in app_list:
                disabled_app_list.append(row["app_name"])

        return disabled_app_list

    def get_enabled_app_list(self, app_type, com_app):
        #TODO : replace with get_app_list()
        """
        Returns enabled app list for given app_type and com_app
        :param app_type: TransitionAppType
        :param com_app : Specify com_app.
        :return: enabled app list, None if app_type and/or com_app doesn't exist
        """
        enabled_app_list = None
        if app_type in self.app_type_list and com_app in self.get_com_app_list():
            app_list = self.get_app_list(app_type=app_type, com_app=com_app, enabled=True)
            if app_list is None:
                app_list = list()

            enabled_app_list = list()
            for row in app_list:
                enabled_app_list.append(row["app_name"])

        return enabled_app_list

    def get_available_app_list(self, app_type, com_app=None):
        """
        Returns available app list for given app_type
        :param app_type: TransitionAppType
        :param com_app: If specified, only return available apps for specified com_app
        :return: available app list
        :raise: ConfigurationException() if com_app parameter is not None or Tuple()
        """
        cursor = self._sqlite.cursor()
        if com_app is None:
            cursor.execute(SQL_SELECT_APP_LIST_BY_TYPE, (app_type, ))
        elif isinstance(com_app, tuple):
            com_app_tuple = ''
            for app in com_app:
                if len(com_app_tuple) == 0:
                    com_app_tuple += "'" + app + "'"
                else:
                    com_app_tuple += ", '" + app + "'"

            cursor.execute(SQL_SELECT_APP_LIST_BY_TYPE_AND_COM_APP.format(com_app_tuple), (app_type,))
        else:
            mesg = "get_available_app_list() : com_app parameter must be None or Tuple(). Get " + repr(com_app)
            print(mesg)
            raise ConfigurationException(mesg)

        available_app_list = list()

        for row in cursor:
            available_app_list.append(row["name"])

        return available_app_list

    def get_app_list(self, app_type=None, app_name=None, com_app=None, enabled=None):
        """
        Return a filtered row[] list() of available apps.
        :param app_type: filter list on given app_type. Default is None: don't filter
        :param app_name: filter list on given app_name. Default is None: don't filter
        :param com_app: filter list on given com_app. Default is None: don't filter
        :param enabled: filter list on enabled flag. Default is None: don't filter
        :return: list of row. None if no records are found
        """
        cursor = self._sqlite.cursor()
        cursor.execute(SQL_SELECT_APP_WORKS_WITH_COM_APP)

        app_list = list()
        for row in cursor:
            app_list.append(row)
            if ((app_type is not None and app_type != row['app_type'])
                or (app_name is not None and app_name != row['app_name'])
                or (com_app is not None and com_app != row['com_app'])
                or (enabled is not None and enabled != row['enabled'])):
                app_list.remove(row)

        if len(app_list) == 0:
            app_list = None

        return app_list

    def _get_app_type_id(self, app_type):
        """Return app_type_id from its name"""
        cursor = self._sqlite.cursor()
        cursor.execute(SQL_SELECT_APP_TYPE_ID, (app_type, ))
        row = cursor.fetchone()
        if row is not None:
            app_id = row["rowid"]
        else:
            app_id = -1
        return app_id

    def _get_com_app_type_id(self, com_app):
        """Return com_app_id from its name
        :param com_app: com application name
        :return: com application id in db
        """
        cursor = self._sqlite.cursor()
        cursor.execute(SQL_SELECT_COM_APP, (com_app, ))
        row = cursor.fetchone()
        if row is not None:
            app_id = row["rowid"]
        else:
            app_id = -1
        return app_id

    def _get_app_id(self, app_type, app_name):
        """Return app_id from its name and type
        :param app_type: type of application
        :param app_name: name of application
        :return: application id in db or -1 if app doesn't exist in db
        """
        cursor = self._sqlite.cursor()
        cursor.execute(SQL_SELECT_APP_ID, (app_type, app_name))
        row = cursor.fetchone()
        if row is None:
            app_id = -1
        else:
            app_id = row["rowid"]
        return app_id


    @staticmethod
    def _get_app_info_from_desc(desc):
        """Extract author and version from app_description
        :param desc: application description
        """
        import re

        author = ''
        version = ''

        if isinstance(desc, str):
            matches = re.finditer("Author:\s+(.+)+\n", desc, re.M)
            print(repr(matches))
            for m in matches:
                if m is not None:
                    if author == '':
                        author = m.group(1)
                    else:
                        author += '\n' + m.group(1)

            m = re.search("Version:\s+(.+)\n", desc)
            if m is not None:
                version = m.group(1)

        return author, version

    @staticmethod
    def _app_import(app_type, app_name):
        """return imported module
        :param app_type: package
        :param app_name: sub-package
        :return: module or None on error
        """
        module = None
        try:
            module = inspect.importlib.import_module("{}.{}".format(app_type, app_name))
        except ImportError as ie:
            print("_app_import() : ERROR : can't import {}.{}. {} is not registerable".format(app_type, app_name,
                                                                                              app_name))
            print("_app_import() :", ie.msg)
        except (SyntaxError, IndentationError, TabError) as se:
            #https://docs.python.org/3.4/library/exceptions.html#exception-hierarchy
            print("_app_import() : ERROR : can't import {}.{}. {} is not registerable".format(app_type, app_name,
                                                                                              app_name))
            print("_app_import() :", se.msg, se.filename, se.lineno, se.text)
            print("*** NAUGHTY PROGRAMMER!!!")
            print("*** SPANK SPANK SPANK!!!")
            print("*** Now go fix your code. Tut tut tut!")

        return module

    def _app_add(self, app_type, app_name, fire_event=True):
        """
        add an application and its info to db. Status is disabled.
        :param app_type: app type (docapp, addin)
        :param app_name: app name
        :param fire_event: Can be set to False to add app silently
        :return: app_id or -1 on error
        """
        print("_app_add()", app_type, app_name)
        app_path = self.get_app_type_path(app_type) + "\\" + app_name
        ret_val = -1

        if self._get_app_id(app_type, app_name) < 0:

            module = self._app_import(app_type, app_name)
            if module is not None:
                app_desc = self._get_app_desc(app_type, app_name)
                author, version = Configuration._get_app_info_from_desc(app_desc)
                digest = self._generate_app_digests(app_path)

                works_with = module.com_app
                app_type_id = self._get_app_type_id(app_type)
                if app_type_id > 0:
                    cursor = self._sqlite.cursor()
                    print(SQL_INSERT_APP)
                    cursor.execute(SQL_INSERT_APP, (app_name,
                                                    author,
                                                    version,
                                                    app_desc,
                                                    app_path,
                                                    app_type_id,
                                                    digest["SHA256"],
                                                    digest["SHA512"],
                                                    digest["WHIRLPOOL"]))

                    #Insert links with com_app
                    self._sqlite.commit()
                    app_id = self._get_app_id(app_type, app_name)
                    for com_app in works_with:
                        com_app_id = self._get_com_app_type_id(com_app)
                        if com_app_id > 0:
                            cursor.execute(SQL_INSERT_APP_WORKS_WITH_COM_APP, (app_id, com_app_id, False))
                        else:
                            #display warning
                            print("_app_add() : WARNING : ignore com_app link. com_app '{}' is unknown !"
                                  .format(com_app))

                    self._sqlite.commit()

                    if fire_event:
                        self._fire_event("on_app_add", (app_type, app_name))

                    ret_val = self._get_app_id(app_type, app_name)
            else:
                print("_app_add() : module is not importable ! see previous messages.")
        else:
            print("_app_add() : ERROR : app_type {} is unknown !".format(app_type))
            print("_app_add() : ERROR : {} is not registerable".format(app_name))

        return ret_val

    def _app_del(self, app_type, app_name, fire_event=True):
        """
        Remove application from db
        :param app_type: application type
        :param app_name: application name
        :param fire_event: can be set to False to remove silently
        :return: app_id of removed app
        """

        print("_app_del()", app_type, app_name)
        app_id = self._get_app_id(app_type, app_name)
        if app_id > 0:
            cursor = self._sqlite.cursor()
            cursor.execute(SQL_DELETE_APP_WORKS_WITH_COM_APP_BY_ID, (app_id, ))
            cursor.execute(SQL_DELETE_APP_BY_ID, (app_id, ))
            self._sqlite.commit()

            if fire_event:
                self._fire_event("on_app_del", (app_type, app_name))
        else:
            app_id = -1

        return app_id

    def _app_update(self, app_type, app_name, digests, fire_event=True):
        """
        Update app info fields and work_with association.

        app fields are :
            author TEXT,
            version TEXT,
            description TEXT,
            SHA256 VARCHAR(64),
            SHA512 VARCHAR(128),
            WHIRLPOOL VARCHAR(128)

        :param app_type
        :param app_name
        :param digests
        :param fire_event
        :return app_id or -1 if app doesn't exists
        """

        app_id = self._get_app_id(app_type, app_name)
        if app_id > 0:
            module = self._app_import(app_type, app_name)
            if module is not None:
                app_desc = self._get_app_desc(app_type, app_name)
                author, version = Configuration._get_app_info_from_desc(app_desc)
                works_with = module.com_app

                cursor = self._sqlite.cursor()
                print(SQL_INSERT_APP)
                cursor.execute(SQL_UPDATE_APP, (author,
                                                version,
                                                app_desc,
                                                digests["SHA256"],
                                                digests["SHA512"],
                                                digests["WHIRLPOOL"],
                                                app_id))

                #Insert relations with com_app
                cursor.execute(SQL_DELETE_APP_WORKS_WITH_COM_APP_BY_ID, (app_id, ))

                for com_app in works_with:
                    com_app_id = self._get_com_app_type_id(com_app)
                    if com_app_id > 0:
                        cursor.execute(SQL_INSERT_APP_WORKS_WITH_COM_APP, (app_id, com_app_id, False))
                    else:
                        #display warning
                        print("_app_update() : WARNING : ignore com_app link. com_app '{}' is unknown !"
                              .format(com_app))

                self._sqlite.commit()

                if fire_event:
                    self._fire_event("on_app_update", (app_type, app_name))
            else:
                print("_app_update() : module is not importable ! see previous messages.")
        return app_id

    def get_app_mode(self, app_type, app_name):
        """
        Check application mode (if Manifest file is available => usermode)
        :param app_type:
        :param app_name:
        :return: devmode/usermode or None if app_type and/or app_name doesn't exist
        """
        app_info = self.get_app_info(app_type, app_name)
        mode = None
        if app_info is not None:
            if os.path.exists(app_info['path'] + "\\Manifest"):
                mode = 'usermode'
            else:
                mode = 'devmode'

        return mode

    def get_app_state(self, app_type, app_name):
        """
        get registered app state (from digests)
        :param app_type:
        :param app_name:
        :return: 0 : unchanged, 1 : updated, -1 corrupted, None if app_type and/or app_name doesn't exist
        """
        app_info = self.get_app_info(app_type, app_name)
        if app_info is None:
            ret_val = None
        else:
            ret_val = 0
            if self.get_app_mode(app_type, app_name) == 'usermode':
                # check Manifest against generated digests
                digest_ok, _ = self._verify_manifest_digests(app_info['path'])
                if not digest_ok:
                    ret_val = -1
                else:
                    digest_ok, _ = self._verify_db_digests(app_info['path'])
                    if not digest_ok:
                        ret_val = 1
            else:
                digest_ok, _ = self._verify_db_digests(app_info['path'])
                if not digest_ok:
                    ret_val = 1

        return ret_val

    def print_app_list(self, app_type=None):
        """
        Displays names, state (loaded or not) and descriptions of the available
        docapp.
        :param app_type: application type. Default None for all registered app_type
        """

        if app_type is None:
            for app_type in self.app_type_list:
                self.print_app_list(app_type)
        else:
            print("\nAvailable {} :".format(app_type))

            app_list = self.get_app_list(app_type)
            if app_list is not None:
                name = ''
                for item in app_list:
                    # print name and description once
                    if name != item['app_name']:
                        print("\n" + item['description'])
                        print(item['app_name'], "is available for :")
                        name = item['app_name']

                    # print status for each com_app
                    if item['enabled']:
                        print("* {} [ENABLED]".format(item['com_app']))
                    else:
                        print("* {} [DISABLED]".format(item['com_app']))
            else:
                print("No application registered for", app_type)
                print("Available application types are", self.app_type_list)

    @staticmethod
    def _get_app_desc(app_type, app_name):
        """
        Return the description of the given app from file system
        :param app_type
        :param app_name
        :returns module description
        """
        desc = ""

        # Dynamic import of the package - to be able to load comments
        module = Configuration._app_import(app_type, app_name)

        if module is not None:
            # return top comments of the package
            desc = inspect.getcomments(module)

        return desc


if __name__ == "__main__":
    config = Configuration()
    config.reset()
    m = config._sqlite
    # for app_type in config.app_type_list:
    #     print(config.get_available_app_list(app_type))