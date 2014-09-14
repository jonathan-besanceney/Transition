# ------------------------------------------------------------------------------
# Name:        configurationtest
# Purpose:     Configuration Class unit tests
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

import unittest
import sys
import os
from transitioncore.configuration import Configuration
from transitioncore.eventsinterface.configeventinterface import ConfigEventsInterface

sys.path.append(os.path.abspath(os.path.dirname(__file__)))

class ConfigEvt(ConfigEventsInterface):
    def __init__(self):
        super(ConfigEvt, self).__init__()
        self.last_fired_evt = ''
        self.last_fired_evt_args = dict()
        self.event_list = ('on_app_add', 'on_app_del', 'on_app_disable', 'on_app_enable', 'on_app_update')

    def on_app_add(self, app_type, app_name):
        self.last_fired_evt = 'on_app_add'
        self.last_fired_evt_args = {'app_type': app_type, 'app_name': app_name}

    def on_app_del(self, app_type, app_name):
        self.last_fired_evt = 'on_app_del'
        self.last_fired_evt_args = {'app_type': app_type, 'app_name': app_name}

    def on_app_disable(self, app_type, app_name):
        self.last_fired_evt = 'on_app_disable'
        self.last_fired_evt_args = {'app_type': app_type, 'app_name': app_name}

    def on_app_enable(self, app_type, app_name):
        self.last_fired_evt = 'on_app_enable'
        self.last_fired_evt_args = {'app_type': app_type, 'app_name': app_name}

    def on_app_update(self, app_type, app_name):
        self.last_fired_evt = 'on_app_update'
        self.last_fired_evt_args = {'app_type': app_type, 'app_name': app_name}

class ConfigurationTest(unittest.TestCase):

    def setUp(self):
        import shutil
        import tests
        self.test_dir = tests.__path__[0] + '\\test_files'
        self.config = Configuration()
        self.config.reset()
        self.evt = ConfigEvt()
        self.config.add_event_listener(self.evt)
        try:
            shutil.rmtree(self.test_dir)
        except Exception as e:
            print("It shouldn't be a problem : ", repr(e), self.test_dir)

    def tearDown(self):
        import shutil
        self.config = None
        try:
            shutil.rmtree(self.test_dir)
        except Exception as e:
            print(repr(e))

    def test__app_add(self):
        #Can't insert if app exists, so we delete it before
        self.config._app_del('complugin', 'config')

        #cleanup
        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()

        app_id = self.config._app_add('complugin', 'config')
        self.assertLess(0, app_id)

        #check if "work with" records are ok
        cursor = self.config._sqlite_cnx.cursor()
        cursor.execute("""SELECT short_name
                            FROM com_app
                                INNER JOIN app_works_with_com_app ON com_app.rowid = app_works_with_com_app.id_com_app
                            WHERE app_works_with_com_app.id_app = ?
                            """, (app_id, ))

        import complugin.config
        com_app_count = 0
        for row in cursor:
            com_app_count += 1
            self.assertIn(row["short_name"], complugin.config.com_app)

        self.assertEqual(len(complugin.config.com_app), com_app_count)

        #event must be fired
        self.assertEqual('on_app_add', self.evt.last_fired_evt)
        self.assertDictEqual({'app_type': 'complugin', 'app_name': 'config'}, self.evt.last_fired_evt_args)

        #cleanup
        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()

        app_id = self.config._app_add('complugin', 'config')
        self.assertEqual( -1, app_id)
        self.assertEqual('', self.evt.last_fired_evt)
        self.assertDictEqual({}, self.evt.last_fired_evt_args)

    def test__app_del(self):
        #cleanup
        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()

        #must work
        app_id = self.config._app_del('complugin', 'config')
        self.assertGreater(app_id, 0)
        #event must be fired
        self.assertEqual('on_app_del', self.evt.last_fired_evt)
        self.assertDictEqual({'app_type': 'complugin', 'app_name': 'config'}, self.evt.last_fired_evt_args)

        #cleanup
        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()

        #must return an error
        app_id = self.config._app_del('complugin', 'config')
        self.assertEqual(-1, app_id)
        self.assertEqual('', self.evt.last_fired_evt)
        self.assertDictEqual({}, self.evt.last_fired_evt_args)

    def test__app_update(self):
        """
        app fields are :
            author TEXT,
            version TEXT,
            description TEXT,
            SHA256 VARCHAR(64),
            SHA512 VARCHAR(128),
            WHIRLPOOL VARCHAR(128)
        """

        import complugin.config
        #reset all updatable fields
        app_id = self.config._get_app_id('complugin', 'config')
        cursor = self.config._sqlite_cnx.cursor()
        cursor.execute("""UPDATE app
                        SET author='', version='', description='', SHA256='', SHA512='', WHIRLPOOL=''
                        WHERE rowid = ?""", (app_id, ))
        cursor.execute("""DELETE FROM app_works_with_com_app
                        WHERE id_app = ?""", (app_id, ))
        self.config._sqlite_cnx.commit()

        #reset events
        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()

        #get expected values
        description = self.config._get_app_desc('complugin', 'config')
        author, version = self.config._get_app_info_from_desc(description)
        digests = self.config._generate_app_digests(complugin.config.__path__[0])
        works_with = complugin.config.com_app

        #update config
        ret_val = self.config._app_update('complugin', 'config', digests)

        #check ret_val
        self.assertEqual(app_id, ret_val)

        #check written values
        cursor = self.config._sqlite_cnx.cursor()
        cursor.execute("""SELECT * FROM app
                WHERE rowid = ?""", (app_id, ))
        infos = cursor.fetchone()
        self.assertEqual(description, infos['description'])
        self.assertEqual(author, infos['author'])
        self.assertEqual(version, infos['version'])
        self.assertEqual(digests['SHA256'], infos['SHA256'])
        self.assertEqual(digests['SHA512'], infos['SHA512'])
        self.assertEqual(digests['WHIRLPOOL'], infos['WHIRLPOOL'])

        cursor = self.config._sqlite_cnx.cursor()
        cursor.execute("""SELECT short_name
                FROM app_works_with_com_app
                    INNER JOIN com_app ON app_works_with_com_app.id_com_app = com_app.rowid
                WHERE id_app = ?""", (app_id, ))

        com_app_count = 0
        for row in cursor:
            com_app_count += 1
            self.assertIn(row['short_name'], works_with)

        self.assertEqual(len(works_with), com_app_count)

        #check event fired
        self.assertEqual('on_app_update', self.evt.last_fired_evt)
        self.assertDictEqual({'app_type': 'complugin', 'app_name': 'config'}, self.evt.last_fired_evt_args)

        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()

        #check error
        ret_val = self.config._app_update('com', 'config', digests)
        self.assertEqual(ret_val, -1)

        #check event fired
        self.assertEqual('', self.evt.last_fired_evt)
        self.assertDictEqual({}, self.evt.last_fired_evt_args)

    def test__create_tables(self):
        ##Test setup

        sqlite_cnx = self.config._sqlite
        #no events is fired at this stage
        self.assertEqual('', self.evt.last_fired_evt)
        self.assertDictEqual({}, self.evt.last_fired_evt_args)

        #must drop all tables to test create_tables
        cursor = sqlite_cnx.cursor()
        cursor.execute("SELECT tbl_name FROM sqlite_master WHERE type = 'table'")
        expected_table_list = list()
        for row in cursor:
            print("table =>", row["tbl_name"])
            expected_table_list.append(row["tbl_name"])

        for table in expected_table_list:
            sqlite_cnx.cursor().execute("DROP TABLE {}".format(table))
        sqlite_cnx.commit()

        #db must be empty
        cursor = sqlite_cnx.cursor().execute("SELECT tbl_name FROM sqlite_master WHERE type = 'table'")
        self.assertIsNone(cursor.fetchone())

        ##test

        #tables creation
        self.config._create_tables()

        #Check that we have same tables as before
        cursor = sqlite_cnx.cursor().execute("SELECT tbl_name FROM sqlite_master WHERE type = 'table'")
        table_count = 0
        for row in cursor:
            table_count += 1
            self.assertIn(row['tbl_name'], expected_table_list)

        #same number as defined in config
        self.assertEqual(len(self.config.create_tables), table_count)

        #Check we have initial values

        # look for app_types
        cursor = sqlite_cnx.cursor().execute("SELECT name FROM app_type")
        app_type_count = 0
        for app_type in cursor:
            app_type_count += 1
            self.assertIn(app_type['name'], self.config.app_types)

        #same number as defined in config
        self.assertEqual(len(self.config.app_types), app_type_count)

        #look for com_apps, in lower case
        com_apps = list()
        for com_app in self.config.com_apps:
            com_apps.append(com_app.lower())

        cursor = sqlite_cnx.cursor().execute("SELECT short_name FROM com_app")
        com_app_count = 0
        for com_app in cursor:
            com_app_count += 1
            self.assertIn(com_app['short_name'], com_apps)

        #same number as defined in config
        self.assertEqual(len(self.config.com_apps), com_app_count)

        #control that no events was fired during db creation
        self.assertEqual('', self.evt.last_fired_evt)
        self.assertDictEqual({}, self.evt.last_fired_evt_args)

    def test__get_app_id(self):
        app_id = self.config._get_app_id('complugin', 'config')
        cursor = self.config._sqlite_cnx.cursor()
        cursor.execute("SELECT rowid FROM app WHERE name = ?", ('config', ))
        expected_app_id = cursor.fetchone()['rowid']
        self.assertEqual(expected_app_id, app_id)

        app_id = self.config._get_app_id('complugi', 'config')
        self.assertEqual(-1, app_id)


    def test__get_app_desc(self):
        import inspect
        import complugin.config
        expected_desc = inspect.getcomments(complugin.config)
        description = self.config._get_app_desc('complugin', 'config')
        self.assertEqual(expected_desc, description)

        #check event fired
        self.assertEqual('', self.evt.last_fired_evt)
        self.assertDictEqual({}, self.evt.last_fired_evt_args)

        description = self.config._get_app_desc('com', 'config')
        self.assertEqual(-1, description)

    def test__get_app_info_from_desc(self):
        description = """
        #
        # Author: Jonathan
        # Author: Besanceney
        #
        # Version: 1.0
        # Version: 2.0
        """
        expected_author = "Jonathan\nBesanceney"
        expected_version = "1.0"
        author, version = self.config._get_app_info_from_desc(description)
        self.assertEqual(expected_author, author)
        self.assertEqual(expected_version, version)

    def test__generate_app_digests(self):
        import os
        import inspect

        os.mkdir(self.test_dir)

        #Empty dir, no digest
        digests = self.config._generate_app_digests(self.test_dir)
        self.assertIsNone(digests)

        #wrong directory, no digests
        digests = self.config._generate_app_digests(self.test_dir + "1")
        self.assertIsNone(digests)

        #create a file
        f = open(self.test_dir + '\\__init__.py', 'w')
        f.close()

        #ok with digests we expect
        expected_digests = self.config._generate_app_digests(self.test_dir)
        self.assertIsNotNone(expected_digests)


        #now we modify previous file. Digests must differ.
        with open(self.test_dir + '\\__init__.py', 'w') as f:
            f.write("print('Hello world!')")

        digests = self.config._generate_app_digests(self.test_dir)
        for key in digests.keys():
            self.assertNotEqual(expected_digests[key], digests[key])

    def test__get_sqlite(self):
        import sqlite3
        import os
        sqlite_cnx = self.config._sqlite
        #database must exist now
        self.assertTrue(os.path.exists(os.getenv("USERPROFILE") + "\\.transition.s3db"))
        #no events
        self.assertEqual('', self.evt.last_fired_evt)
        self.assertDictEqual({}, self.evt.last_fired_evt_args)
        #be connected :p
        self.assertIsInstance(sqlite_cnx, sqlite3.Connection)

    def test__get_app_type_id(self):
        import tests
        #prepare test
        app_type = "tests"
        cursor = self.config._sqlite.cursor()
        cursor.execute("INSERT INTO app_type VALUES (?, ?)", (app_type, tests.__path__[0]))
        self.config._sqlite_cnx.commit()
        expected_app_type_id = cursor.lastrowid

        #must work
        app_type_id = self.config._get_app_type_id(app_type)
        self.assertEqual(expected_app_type_id, app_type_id)

        #check for error return
        app_type_id = self.config._get_app_type_id("contraption")
        self.assertEqual(-1, app_type_id)

    def test__get_com_app_type_id(self):
        #prepare test
        app_type = "unknown com app"
        cursor = self.config._sqlite.cursor()
        cursor.execute("INSERT INTO com_app VALUES (?)", (app_type, ))
        self.config._sqlite_cnx.commit()
        expected_app_type_id = cursor.lastrowid

        #must work
        app_type_id = self.config._get_com_app_type_id(app_type)
        self.assertEqual(expected_app_type_id, app_type_id)

        #check for error return
        app_type_id = self.config._get_com_app_type_id("contraption")
        self.assertEqual(-1, app_type_id)

    def test__read_app_manifest(self):
        from transitioncore.exceptions.configurationexception import ConfigurationException
        import inspect
        import pickle

        os.mkdir(self.test_dir)
        with open(self.test_dir + '\\__init__.py', 'w') as f:
            f.write("print('Hello world!')")

        module = inspect.importlib.import_module("tests.test_files")

        #check without Manifest, raise ConfigurationException
        self.assertRaises(ConfigurationException, self.config._read_app_manifest, module.__path__[0] + '\\')

        #check for valid Manifest
        self.config.write_app_manifest(module.__path__[0] + '\\')
        digests = self.config._read_app_manifest(module.__path__[0] + '\\')
        self.assertIsInstance(digests, dict)
        self.assertIn("SHA256", digests.keys())
        self.assertIn("SHA512", digests.keys())
        self.assertIn("WHIRLPOOL", digests.keys())
        self.assertEqual(64, len(digests["SHA256"]))
        self.assertEqual(128, len(digests["SHA512"]))
        self.assertEqual(128, len(digests["WHIRLPOOL"]))

        #check for invalid Manifest, invalid length in WHIRLPOOL hash, raise ConfigurationException
        with open(module.__path__[0] + '\\Manifest', 'rb') as f:
            digests = pickle.load(f)
            digests["WHIRLPOOL"] = digests["WHIRLPOOL"][:127]
        with open(module.__path__[0] + '\\Manifest', 'wb') as f:
            pickle.dump(digests, f)

        self.assertRaises(ConfigurationException, self.config._read_app_manifest, module.__path__[0] + '\\')

        #check for invalid Manifest, invalid key, raise ConfigurationException
        self.config.write_app_manifest(module.__path__[0] + '\\')
        with open(module.__path__[0] + '\\Manifest', 'rb') as f:
            digests = pickle.load(f)

        digests["WHIRLPOO"] = digests["WHIRLPOOL"]
        del digests["WHIRLPOOL"]

        with open(module.__path__[0] + '\\Manifest', 'wb') as f:
            pickle.dump(digests, f)

        self.assertRaises(ConfigurationException, self.config._read_app_manifest, module.__path__[0] + '\\')

        #check with empty Manifest
        os.remove(module.__path__[0] + '\\Manifest')
        f = open(module.__path__[0] + '\\Manifest', 'wb')
        f.close()
        self.assertRaises(ConfigurationException, self.config._read_app_manifest, module.__path__[0] + '\\')

    def test__verify_db_digests(self):
        #Invalid dir, Unknown app
        status, digests = self.config._verify_db_digests(self.test_dir)
        self.assertFalse(status)
        self.assertIsNone(digests)

        #Valid empty dir, Unknown app
        os.mkdir(self.test_dir)
        status, digests = self.config._verify_db_digests(self.test_dir)
        self.assertFalse(status)
        self.assertIsNone(digests)

        #Valid dir with __init__.py, Unknown app
        with open(self.test_dir + '\\__init__.py', 'w') as f:
            f.write("print('Hello world!')")

        status, digests = self.config._verify_db_digests(self.test_dir)
        self.assertFalse(status)
        self.assertIsInstance(digests, dict)
        for key in ('SHA256', 'SHA512', 'WHIRLPOOL'):
            self.assertIsNotNone(digests[key])
            if key == 'SHA256':
                self.assertEqual(64, len(digests[key]))
            else:
                self.assertEqual(128, len(digests[key]))

        #valid Config dir, known app, digests differs
        cursor = self.config._sqlite.cursor()
        cursor.execute("UPDATE app SET SHA256='' WHERE rowid=?", (self.config._get_app_id('complugin', 'config'), ))
        self.config._sqlite_cnx.commit()

        import complugin.config
        status, digests = self.config._verify_db_digests(complugin.config.__path__[0])
        self.assertFalse(status)
        self.assertIsInstance(digests, dict)

        #valid Config dir, known app, digests ok
        #ensure devmode by removing Manifest file if exist
        if os.path.exists(complugin.config.__path__[0] + '\\Manifest'):
            os.remove(complugin.config.__path__[0] + '\\Manifest')

        #update app to have a new manifest stored in db
        self.config._app_update('complugin', 'config', digests)

        status, digests = self.config._verify_db_digests(complugin.config.__path__[0])
        self.assertTrue(status)
        self.assertIsInstance(digests, dict)
        for key in ('SHA256', 'SHA512', 'WHIRLPOOL'):
            self.assertIsNotNone(digests[key])
            if key == 'SHA256':
                self.assertEqual(64, len(digests[key]))
            else:
                self.assertEqual(128, len(digests[key]))

        expected_digests = self.config._generate_app_digests(complugin.config.__path__[0])
        self.assertDictEqual(expected_digests, digests)

    def test__verify_manifest_digests(self):
        #Invalid dir, no Manifest
        #ensure devmode by removing Manifest file if exist
        if os.path.exists(self.test_dir):
            os.remove(self.test_dir)
        status, digests = self.config._verify_manifest_digests(self.test_dir)
        self.assertFalse(status)
        self.assertIsNone(digests)

        #Valid dir without __init__.py, no Manifest
        os.mkdir(self.test_dir)
        status, digests = self.config._verify_manifest_digests(self.test_dir)
        self.assertFalse(status)
        self.assertIsNone(digests)

        #Valid dir with __init__.py, no Manifest
        with open(self.test_dir + '\\__init__.py', 'w') as f:
            f.write("print('Hello world!')")

        status, digests = self.config._verify_manifest_digests(self.test_dir)
        self.assertFalse(status)
        self.assertIsInstance(digests, dict)
        for key in ('SHA256', 'SHA512', 'WHIRLPOOL'):
            self.assertIsNotNone(digests[key])
            if key == 'SHA256':
                self.assertEqual(64, len(digests[key]))
            else:
                self.assertEqual(128, len(digests[key]))

        #Valid dir with __init__.py, Empty Manifest
        f = open(self.test_dir + '\\Manifest', 'w')
        f.close()
        status, digests = self.config._verify_manifest_digests(self.test_dir)
        self.assertFalse(status)
        self.assertIsInstance(digests, dict)
        for key in ('SHA256', 'SHA512', 'WHIRLPOOL'):
            self.assertIsNotNone(digests[key])
            if key == 'SHA256':
                self.assertEqual(64, len(digests[key]))
            else:
                self.assertEqual(128, len(digests[key]))

        #Valid dir with __init__.py, Manifest OK
        expected_digests = self.config.write_app_manifest(self.test_dir)
        status, digests = self.config._verify_manifest_digests(self.test_dir)
        self.assertTrue(status)
        self.assertDictEqual(expected_digests, digests)

        #Valid dir with __init__.py, Manifest mismatch
        import pickle
        with open(self.test_dir + '\\Manifest', 'rb') as f:
            digests = pickle.load(f)
            digests["WHIRLPOOL"] = digests["WHIRLPOOL"][:127]
        with open(self.test_dir + '\\Manifest', 'wb') as f:
            pickle.dump(digests, f)

        status, digests = self.config._verify_manifest_digests(self.test_dir)
        self.assertFalse(status)

    def test_todo_disable_app(self):
        self.assertTrue(False)

    def test_todo_enable_app(self):
        self.assertTrue(False)

    def test_get_app_info(self):
        import complugin.config
        app_info1 = self.config.get_app_info('complugin', 'config')
        self.assertIsNotNone(app_info1)
        app_info2 = self.config.get_app_info_by_path(complugin.config.__path__[0])
        self.assertIsNotNone(app_info2)

        for field in app_info1.keys():
            self.assertEqual(app_info1[field], app_info2[field])

    def test_get_app_info_by_path(self):
        #with unknown path
        self.assertIsNone(self.config.get_app_info_by_path(self.test_dir))

        #with known path
        import complugin.config
        import sqlite3
        app_infos = self.config.get_app_info_by_path(complugin.config.__path__[0])
        self.assertIsInstance(app_infos, sqlite3.Row)

    def test_todo_get_app_list(self):
        self.assertTrue(False)

    def test_todo_get_app_mode(self):
        self.assertTrue(False)

    def test_todo_get_app_state(self):
        self.assertTrue(False)

    def test_todo_get_app_status(self):
        self.assertTrue(False)

    def test_todo_get_app_type_path(self):
        self.assertTrue(False)

    def test_todo_get_available_app_list(self):
        self.assertTrue(False)

    def test_todo_get_disabled_app_list(self):
        self.assertTrue(False)

    def test_todo_get_enabled_app_list(self):
        self.assertTrue(False)

    def test_todo_print_app_list(self):
        self.assertTrue(False)

    def test_reset(self):
        #reset is already called in setUp
        #config db must not exist
        import os
        self.assertFalse(os.path.exists(os.getenv("USERPROFILE") + "\\.transition.s3db"))

    def test_todo_update_inventory(self):
        self.assertTrue(False)

    def test_write_app_manifest(self):
        # from transitioncore.exceptions.configurationexception import ConfigurationException

        #check with invalid path
        self.assertIsNone(self.config.write_app_manifest(self.test_dir))

        #check with empty path
        os.mkdir(self.test_dir)
        self.assertIsNone(self.config.write_app_manifest(self.test_dir))

        #check with valid app,
        with open(self.test_dir + '\\__init__.py', 'w') as f:
            f.write("print('Hello world!')")

        self.assertIsInstance(self.config.write_app_manifest(self.test_dir), dict)

        #TODO : find a way to raise ConfigurationException (IOError)

if __name__ == '__main__':
    unittest.main()