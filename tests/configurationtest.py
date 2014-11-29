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
        print('fired event', self.last_fired_evt, self.last_fired_evt_args)

    def on_app_del(self, app_type, app_name):
        self.last_fired_evt = 'on_app_del'
        self.last_fired_evt_args = {'app_type': app_type, 'app_name': app_name}
        print('fired event', self.last_fired_evt, self.last_fired_evt_args)

    def on_app_disable(self, app_type, app_name, com_app_tuple):
        self.last_fired_evt = 'on_app_disable'
        self.last_fired_evt_args = {'app_type': app_type, 'app_name': app_name, 'com_app_tuple': com_app_tuple}
        print('fired event', self.last_fired_evt, self.last_fired_evt_args)

    def on_app_enable(self, app_type, app_name, com_app_tuple):
        self.last_fired_evt = 'on_app_enable'
        self.last_fired_evt_args = {'app_type': app_type, 'app_name': app_name, 'com_app_tuple': com_app_tuple}
        print('fired event', self.last_fired_evt, self.last_fired_evt_args)

    def on_app_update(self, app_type, app_name):
        self.last_fired_evt = 'on_app_update'
        self.last_fired_evt_args = {'app_type': app_type, 'app_name': app_name}
        print('fired event', self.last_fired_evt, self.last_fired_evt_args)


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
        self.config._app_del('addin', 'config')

        #cleanup
        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()

        app_id = self.config._app_add('addin', 'config')
        self.assertLess(0, app_id)

        #check if "work with" records are ok
        cnx = self.config._sqlite
        cursor = cnx.cursor()
        cursor.execute("""SELECT short_name
                            FROM com_app
                                INNER JOIN app_works_with_com_app ON com_app.rowid = app_works_with_com_app.id_com_app
                            WHERE app_works_with_com_app.id_app = ?
                            """, (app_id, ))

        import addin.config
        com_app_count = 0
        for row in cursor:
            com_app_count += 1
            self.assertIn(row["short_name"], addin.config.com_app)

        cursor.close()
        self.config._close_sqlite()
        self.assertEqual(len(addin.config.com_app), com_app_count)

        #event must be fired
        self.assertEqual('on_app_add', self.evt.last_fired_evt)
        self.assertDictEqual({'app_type': 'addin', 'app_name': 'config'}, self.evt.last_fired_evt_args)

        #cleanup
        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()

        app_id = self.config._app_add('addin', 'config')
        self.assertEqual( -1, app_id)
        self.assertEqual('', self.evt.last_fired_evt)
        self.assertDictEqual({}, self.evt.last_fired_evt_args)

        #TODO : import fake dirty app

    def test__app_del(self):
        #cleanup
        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()

        #must work
        app_id = self.config._app_del('addin', 'config')
        self.assertGreater(app_id, 0)
        #event must be fired
        self.assertEqual('on_app_del', self.evt.last_fired_evt)
        self.assertDictEqual({'app_type': 'addin', 'app_name': 'config'}, self.evt.last_fired_evt_args)

        #cleanup
        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()

        #must return an error
        app_id = self.config._app_del('addin', 'config')
        self.assertEqual(-1, app_id)
        self.assertEqual('', self.evt.last_fired_evt)
        self.assertDictEqual({}, self.evt.last_fired_evt_args)

    def test__app_import(self):
        os.mkdir(self.test_dir)

        #make a fake app with indent problem
        with open(self.test_dir + "\\__init__.py", "w") as f:
            f.write("""
class FakeApp():
    def run():
    pass

app_class = FakeApp
com_app = ('excel', 'unknown com app')
app_type = 'tests'
""")

        self.assertIsNone(self.config._app_import("tests", "test_files"))

        #make a fake app with indent & syntax problem
        with open(self.test_dir + "\\__init__.py", "w") as f:
            f.write("""
class FakeApp():
    def run():
     pass(

app_class = FakeApp
com_app = ('excel', 'unknown com app')
app_type = 'tests'
""")

        self.assertIsNone(self.config._app_import("tests", "test_files"))

        #make a fake app with indent & syntax problem
        with open(self.test_dir + "\\__init__.py", "w") as f:
            f.write("""
class FakeApp():
    def run():
     pass(

app_class = FakeApp
com_app = ('excel', 'unknown com app')
app_type = 'tests'
""")

        self.assertIsNone(self.config._app_import("tests", "test_files"))

        #make a fake app with indent & syntax problem
        with open(self.test_dir + "\\__init__.py", "w") as f:
            f.write("""
class FakeApp():
    def run():
        pass

app_class = FakeApp
com_app = ('excel', 'unknown com app')
app_type = 'tests'

print("Hello World")
""")

        self.assertIsNotNone(self.config._app_import("tests", "test_files"))

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

        import addin.config
        #reset all updatable fields
        app_id = self.config._get_app_id('addin', 'config')
        cnx = self.config._sqlite
        cursor = cnx.cursor()
        cursor.execute("""UPDATE app
                        SET author='', version='', description='', SHA256='', SHA512='', WHIRLPOOL=''
                        WHERE rowid = ?""", (app_id, ))
        cursor.execute("""DELETE FROM app_works_with_com_app
                        WHERE id_app = ?""", (app_id, ))
        cnx.commit()
        self.config._close_sqlite()
        #reset events
        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()

        #get expected values
        description = self.config._get_app_desc('addin', 'config')
        author, version = self.config._get_app_info_from_desc(description)
        digests = self.config._generate_app_digests(addin.config.__path__[0])
        works_with = addin.config.com_app

        #update config
        ret_val = self.config._app_update('addin', 'config', digests)

        #check ret_val
        self.assertEqual(app_id, ret_val)

        #check written values
        cnx = self.config._sqlite
        cursor = cnx.cursor()
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

        cursor.close
        self.config._close_sqlite()
        self.assertEqual(len(works_with), com_app_count)

        #check event fired
        self.assertEqual('on_app_update', self.evt.last_fired_evt)
        self.assertDictEqual({'app_type': 'addin', 'app_name': 'config'}, self.evt.last_fired_evt_args)

        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()

        #check error
        ret_val = self.config._app_update('com', 'config', digests)
        self.assertEqual(ret_val, -1)

        #check event fired
        self.assertEqual('', self.evt.last_fired_evt)
        self.assertDictEqual({}, self.evt.last_fired_evt_args)

        #TODO : import fake dirty app

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
            self.assertIn(app_type['name'], self.config.app_type_list)

        #same number as defined in config
        self.assertEqual(len(self.config.app_type_list), app_type_count)

        #look for com_apps, in lower case
        com_apps = list()
        for com_app in self.config.com_app_list:
            com_apps.append(com_app.lower())

        cursor = sqlite_cnx.cursor().execute("SELECT short_name FROM com_app")
        com_app_count = 0
        for com_app in cursor:
            com_app_count += 1
            self.assertIn(com_app['short_name'], com_apps)

        #same number as defined in config
        self.assertEqual(len(self.config.com_app_list), com_app_count)

        #control that no events was fired during db creation
        self.assertEqual('', self.evt.last_fired_evt)
        self.assertDictEqual({}, self.evt.last_fired_evt_args)

    def test__get_app_id(self):
        app_id = self.config._get_app_id('addin', 'config')
        cnx = self.config._sqlite
        cursor = cnx.cursor()
        cursor.execute("SELECT rowid FROM app WHERE name = ?", ('config', ))
        expected_app_id = cursor.fetchone()['rowid']
        self.assertEqual(expected_app_id, app_id)
        cursor.close()
        self.config._close_sqlite()
        app_id = self.config._get_app_id('complugi', 'config')
        self.assertEqual(-1, app_id)


    def test__get_app_desc(self):
        import inspect
        import addin.config
        expected_desc = inspect.getcomments(addin.config)
        description = self.config._get_app_desc('addin', 'config')
        self.assertEqual(expected_desc, description)

        #check event fired
        self.assertEqual('', self.evt.last_fired_evt)
        self.assertDictEqual({}, self.evt.last_fired_evt_args)

        description = self.config._get_app_desc('com', 'config')
        self.assertEqual("", description)

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

        #check with another app
        config_path = self.config.get_app_info('addin', 'config')['path']
        expected_digests = self.config._generate_app_digests(config_path)
        digests = self.config._generate_app_digests(config_path)
        for key in digests.keys():
            self.assertEqual(expected_digests[key], digests[key])

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
        self.assertEqual(self.config._sqlite_cnx_usage_count, 1)
        self.assertIsInstance(sqlite_cnx, sqlite3.Connection)
        #close cnx
        self.config._close_sqlite()
        self.assertIsNone(self.config._sqlite_cnx)
        self.assertEqual(self.config._sqlite_cnx_usage_count, 0)

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
        cursor.execute("UPDATE app SET SHA256='' WHERE rowid=?", (self.config._get_app_id('addin', 'config'), ))
        self.config._sqlite_cnx.commit()

        import addin.config
        status, digests = self.config._verify_db_digests(addin.config.__path__[0])
        self.assertFalse(status)
        self.assertIsInstance(digests, dict)

        #valid Config dir, known app, digests ok
        #ensure devmode by removing Manifest file if exist
        if os.path.exists(addin.config.__path__[0] + '\\Manifest'):
            os.remove(addin.config.__path__[0] + '\\Manifest')

        #update app to have a new manifest stored in db
        self.config._app_update('addin', 'config', digests)

        status, digests = self.config._verify_db_digests(addin.config.__path__[0])
        self.assertTrue(status)
        self.assertIsInstance(digests, dict)
        for key in ('SHA256', 'SHA512', 'WHIRLPOOL'):
            self.assertIsNotNone(digests[key])
            if key == 'SHA256':
                self.assertEqual(64, len(digests[key]))
            else:
                self.assertEqual(128, len(digests[key]))

        expected_digests = self.config._generate_app_digests(addin.config.__path__[0])
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

    def test_disable_app(self):
        #prepare test

        #enable config for all com_app
        self.assertTrue(self.config.enable_app('addin', 'config'))
        #evt cleanup
        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()

        self.assertTrue(self.config.disable_app('addin', 'config'))
        #event on_enable_app must be trigered
        import addin.config
        self.assertEqual('on_app_disable', self.evt.last_fired_evt)
        self.assertEqual('addin', self.evt.last_fired_evt_args['app_type'])
        self.assertEqual('config', self.evt.last_fired_evt_args['app_name'])
        self.assertListEqual(list(addin.config.com_app), list(self.evt.last_fired_evt_args['com_app_tuple']))

        #look for all com_app status
        for item in self.config.get_app_list(app_type='addin', app_name='config'):
            self.assertFalse(item['enabled'])

        #prepare enabling config for excel only
        self.assertTrue(self.config.enable_app('addin', 'config'))
        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()

        #enable for excel
        self.assertTrue(self.config.disable_app('addin', 'config', 'excel'))
        #event on_enable_app must be trigered
        self.assertEqual('on_app_disable', self.evt.last_fired_evt)
        self.assertEqual('addin', self.evt.last_fired_evt_args['app_type'])
        self.assertEqual('config', self.evt.last_fired_evt_args['app_name'])
        self.assertListEqual(['excel', ], list(self.evt.last_fired_evt_args['com_app_tuple']))

        #look for all com_app status
        for item in self.config.get_app_list(app_type='addin', app_name='config'):
            if item['com_app'] == 'excel':
                self.assertFalse(item['enabled'])
            else:
                self.assertTrue(item['enabled'])

        #try disabling config for Excel again => must return false and no event
        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()
        self.assertFalse(self.config.disable_app('addin', 'config', 'excel'))
        self.assertEqual('', self.evt.last_fired_evt)
        self.assertDictEqual(dict(), self.evt.last_fired_evt_args)

        #try to disable config for all com_app. must return true, event com_app_tuple arg must not contain excel
        self.assertTrue(self.config.disable_app('addin', 'config'))
        self.assertEqual('on_app_disable', self.evt.last_fired_evt)
        self.assertEqual('addin', self.evt.last_fired_evt_args['app_type'])
        self.assertEqual('config', self.evt.last_fired_evt_args['app_name'])
        expected_com_app_list = list(addin.config.com_app)
        expected_com_app_list.remove('excel')
        self.assertListEqual(expected_com_app_list, list(self.evt.last_fired_evt_args['com_app_tuple']))

        #try with a wrong app and app_type, must raise ConfigurationException
        from transitioncore.exceptions.configurationexception import ConfigurationException
        self.assertRaises(ConfigurationException, self.config.disable_app, 'addin', 'onfig')
        self.assertRaises(ConfigurationException, self.config.disable_app, 'contraption', 'config')

    def test_enable_app(self):
        #evt cleanup
        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()

        #test enable config for all com_app
        self.assertTrue(self.config.enable_app('addin', 'config'))

        #event on_enable_app must be trigered
        import addin.config
        self.assertEqual('on_app_enable', self.evt.last_fired_evt)
        self.assertEqual('addin', self.evt.last_fired_evt_args['app_type'])
        self.assertEqual('config', self.evt.last_fired_evt_args['app_name'])
        self.assertListEqual(list(addin.config.com_app), list(self.evt.last_fired_evt_args['com_app_tuple']))

        #look for all com_app status
        for item in self.config.get_app_list(app_type='addin', app_name='config'):
            self.assertTrue(item['enabled'])

        #prepare enabling config for excel only
        self.assertTrue(self.config.disable_app('addin', 'config'))
        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()

        #enable for excel
        self.assertTrue(self.config.enable_app('addin', 'config', 'excel'))
        #event on_enable_app must be trigered
        self.assertEqual('on_app_enable', self.evt.last_fired_evt)
        self.assertEqual('addin', self.evt.last_fired_evt_args['app_type'])
        self.assertEqual('config', self.evt.last_fired_evt_args['app_name'])
        self.assertListEqual(['excel', ], list(self.evt.last_fired_evt_args['com_app_tuple']))

        #look for all com_app status
        for item in self.config.get_app_list(app_type='addin', app_name='config'):
            if item['com_app'] == 'excel':
                self.assertTrue(item['enabled'])
            else:
                self.assertFalse(item['enabled'])

        #try enable config for Excel again => must return false and no event
        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()
        self.assertFalse(self.config.enable_app('addin', 'config', 'excel'))
        self.assertEqual('', self.evt.last_fired_evt)
        self.assertDictEqual(dict(), self.evt.last_fired_evt_args)

        #try to enable config for all com_app. must return true, event com_app_tuple arg must not contain excel
        self.assertTrue(self.config.enable_app('addin', 'config'))
        self.assertEqual('on_app_enable', self.evt.last_fired_evt)
        self.assertEqual('addin', self.evt.last_fired_evt_args['app_type'])
        self.assertEqual('config', self.evt.last_fired_evt_args['app_name'])
        expected_com_app_list = list(addin.config.com_app)
        expected_com_app_list.remove('excel')
        self.assertListEqual(expected_com_app_list, list(self.evt.last_fired_evt_args['com_app_tuple']))

        #try with a wrong app and app_type, must raise ConfigurationException
        from transitioncore.exceptions.configurationexception import ConfigurationException
        self.assertRaises(ConfigurationException, self.config.enable_app, 'addin', 'onfig')
        self.assertRaises(ConfigurationException, self.config.enable_app, 'contraption', 'config')


    def test_get_app_info(self):
        import addin.config
        app_info1 = self.config.get_app_info('addin', 'config')
        self.assertIsNotNone(app_info1)
        app_info2 = self.config.get_app_info_by_path(addin.config.__path__[0])
        self.assertIsNotNone(app_info2)

        for field in app_info1.keys():
            self.assertEqual(app_info1[field], app_info2[field])

    def test_get_app_info_by_path(self):
        #with unknown path
        self.assertIsNone(self.config.get_app_info_by_path(self.test_dir))

        #with known path
        import addin.config
        import sqlite3
        app_infos = self.config.get_app_info_by_path(addin.config.__path__[0])
        self.assertIsInstance(app_infos, sqlite3.Row)

    def test_get_app_list(self):
        #make fake app, registered with excel only
        from transitioncore.configsql import SQL_INSERT_APP_TYPE
        import tests
        #make a fake app
        os.mkdir(self.test_dir)
        with open(self.test_dir + "\\__init__.py", "w") as f:
            f.write("""
class FakeApp():
    def run():
        pass

app_class = FakeApp
com_app = ('excel', )
app_type = 'tests'
""")

        #insert "tests" app_type
        sqlitecnx = self.config._sqlite
        cursor = sqlitecnx.cursor()
        cursor.execute(SQL_INSERT_APP_TYPE, ("tests", tests.__path__[0]))
        sqlitecnx.commit()
        #insert our test_files fake app
        self.config._app_add("tests", "test_files")

        #test app must be included in app_list associated with excel
        app_list = self.config.get_app_list(app_name='test_files')
        self.assertEqual(1, len(app_list))
        for item in app_list:
            self.assertEqual('test_files', item['app_name'])
            self.assertEqual('tests', item['app_type'])
            self.assertEqual('excel', item['com_app'])
            self.assertEqual(0, item['enabled'])

        app_list = self.config.get_app_list(app_type='tests')
        self.assertEqual(1, len(app_list))
        for item in app_list:
            self.assertEqual('test_files', item['app_name'])
            self.assertEqual('tests', item['app_type'])
            self.assertEqual('excel', item['com_app'])
            self.assertEqual(0, item['enabled'])

        app_list = self.config.get_app_list(app_type='tests', app_name='test_files')
        self.assertEqual(1, len(app_list))
        for item in app_list:
            self.assertEqual('test_files', item['app_name'])
            self.assertEqual('tests', item['app_type'])
            self.assertEqual('excel', item['com_app'])
            self.assertEqual(0, item['enabled'])

        app_list = self.config.get_app_list(app_type='tests', app_name='test_files', com_app='excel')
        self.assertEqual(1, len(app_list))
        for item in app_list:
            self.assertEqual('test_files', item['app_name'])
            self.assertEqual('tests', item['app_type'])
            self.assertEqual('excel', item['com_app'])
            self.assertEqual(0, item['enabled'])

        app_list = self.config.get_app_list(app_type='tests', app_name='test_files', com_app='excel', enabled=False)
        self.assertEqual(1, len(app_list))
        for item in app_list:
            self.assertEqual('test_files', item['app_name'])
            self.assertEqual('tests', item['app_type'])
            self.assertEqual('excel', item['com_app'])
            self.assertEqual(0, item['enabled'])

        app_list = self.config.get_app_list(com_app='excel')
        #look for our fake app in this list
        found_items = 0
        for item in app_list:
            if item['app_name'] == 'test_files':
                found_items += 1

        self.assertEqual(1, found_items)

        app_list = self.config.get_app_list(enabled=False)
        #look for our fake app in this list
        found_items = 0
        for item in app_list:
            if item['app_name'] == 'test_files':
                found_items += 1

        self.assertEqual(1, found_items)

        app_list = self.config.get_app_list(app_name="config")
        #look for registered com_app for config
        import addin.config
        self.assertEqual(len(addin.config.com_app), len(app_list))
        for item in app_list:
            self.assertTrue(item['com_app'] in addin.config.com_app)
            self.assertEqual("config", item['app_name'])

    def test_get_app_mode(self):
        self.assertIsNone(self.config.get_app_mode("tests", "test_files"))

        from transitioncore.configsql import SQL_INSERT_APP_TYPE
        import tests
        #make a fake app
        os.mkdir(self.test_dir)
        with open(self.test_dir + "\\__init__.py", "w") as f:
            f.write("""
class FakeApp():
    def run():
        pass

app_class = FakeApp
com_app = ('excel', )
app_type = 'tests'
""")

        #insert "tests" app_type
        sqlitecnx = self.config._sqlite
        cursor = sqlitecnx.cursor()
        cursor.execute(SQL_INSERT_APP_TYPE, ("tests", tests.__path__[0]))
        sqlitecnx.commit()
        #insert our test_files fake app
        self.config._app_add("tests", "test_files")

        #Without Manifest, fake app is in devmode
        self.assertEqual('devmode', self.config.get_app_mode("tests", "test_files"))

        #With Manifest, app is in usermode
        self.config.write_app_manifest(self.config.get_app_info("tests", "test_files")['path'])
        self.assertEqual('usermode', self.config.get_app_mode("tests", "test_files"))

    def test_get_app_state(self):
        #app doesn't exist
        self.assertIsNone(self.config.get_app_state("tests", "test_files"))

        from transitioncore.configsql import SQL_INSERT_APP_TYPE
        import tests
        #make a fake app
        os.mkdir(self.test_dir)
        with open(self.test_dir + "\\__init__.py", "w") as f:
            f.write("""
class FakeApp():
    def run():
        pass

app_class = FakeApp
com_app = ('excel', )
app_type = 'tests'
""")

        #insert "tests" app_type
        sqlitecnx = self.config._sqlite
        cursor = sqlitecnx.cursor()
        cursor.execute(SQL_INSERT_APP_TYPE, ("tests", tests.__path__[0]))
        sqlitecnx.commit()
        #insert our test_files fake app
        self.config._app_add("tests", "test_files")

        #state must be 0 (unchanged)
        self.assertEqual(0, self.config.get_app_state("tests", "test_files"))

        #we change it in dev mode
        with open(self.test_dir + "\\__init__.py", "w") as f:
            f.write("""
class FakeApp():
    def run():
        print('hello world !')

app_class = FakeApp
com_app = ('excel', )
app_type = 'tests'
""")

        #state must be 1 (updated)
        self.assertEqual(1, self.config.get_app_state("tests", "test_files"))

        self.config._app_update("tests", "test_files", self.config._generate_app_digests(self.test_dir))
        self.assertEqual(0, self.config.get_app_state("tests", "test_files"))

        #In usermode now
        self.config.write_app_manifest(self.test_dir)
        #nothing change
        self.assertEqual(0, self.config.get_app_state("tests", "test_files"))

        #we change app.
        with open(self.test_dir + "\\__init__.py", "w") as f:
            f.write("""
class FakeApp():
    def run():
        print('hello world !')
        print('bonjour tout le monde !')

app_class = FakeApp
com_app = ('excel', )
app_type = 'tests'
""")
        #must return -1 (corrupted)
        self.assertEqual(-1, self.config.get_app_state("tests", "test_files"))

        #generate Manifest
        self.config.write_app_manifest(self.test_dir)

        #must return 1 (updated)
        self.assertEqual(1, self.config.get_app_state("tests", "test_files"))

        #update app
        self.config._app_update("tests", "test_files", self.config._generate_app_digests(self.test_dir))
        self.assertEqual(0, self.config.get_app_state("tests", "test_files"))


    def test_get_app_type_path(self):
        #test with unexisting app_type. Must return None

        self.assertIsNone(self.config.get_app_type_path("tests"))

        #create app_type
        import tests
        from transitioncore.configsql import SQL_INSERT_APP_TYPE
        os.mkdir(self.test_dir)
        #insert "tests" app_type
        sqlitecnx = self.config._sqlite
        cursor = sqlitecnx.cursor()
        cursor.execute(SQL_INSERT_APP_TYPE, ("tests", tests.__path__[0]))
        sqlitecnx.commit()

        self.assertEqual(tests.__path__[0], self.config.get_app_type_path("tests"))

        import addin
        self.assertEqual(addin.__path__[0], self.config.get_app_type_path("addin"))
        import docapp
        self.assertEqual(docapp.__path__[0], self.config.get_app_type_path("docapp"))

    def test_get_available_app_list(self):
        from transitioncore.configsql import SQL_INSERT_APP_TYPE
        import tests
        from transitioncore.exceptions.configurationexception import ConfigurationException
        #wrong type for com_app arg. Raises ConfigurationException
        self.assertRaises(ConfigurationException, self.config.get_available_app_list, "addin", "excel")

        #make a fake app
        os.mkdir(self.test_dir)
        with open(self.test_dir + "\\__init__.py", "w") as f:
            f.write("""
class FakeApp():
    def run():
        pass

app_class = FakeApp
com_app = ('excel', 'unknown com app')
app_type = 'tests'
""")

        #insert "tests" app_type
        sqlitecnx = self.config._sqlite
        cursor = sqlitecnx.cursor()
        cursor.execute(SQL_INSERT_APP_TYPE, ("tests", tests.__path__[0]))
        sqlitecnx.commit()
        #insert our test_files fake app
        self.config._app_add("tests", "test_files")

        #now we can check if get_available_app_list result is valid
        app_list = self.config.get_available_app_list("tests")

        self.assertTrue("test_files" in app_list)
        self.assertTrue("config" not in app_list)

        app_list = self.config.get_available_app_list("tests", ('excel', ))
        self.assertTrue("test_files" in app_list)
        self.assertTrue("config" not in app_list)

        from addin.config import com_app
        app_list = self.config.get_available_app_list("addin", com_app)
        self.assertTrue("test_files" not in app_list)
        self.assertTrue("config" in app_list)

    def test_get_com_app_list(self):
        expectedList = list()
        for com_app in self.config.com_app_list:
            expectedList.append(com_app.lower())
        self.assertListEqual(expectedList, self.config.get_com_app_list())

    def test_get_disabled_app_list(self):
        #app_type tests doesn't exist, must return None
        self.assertIsNone(self.config.get_disabled_app_list("tests", "excel"))

        #all apps are disabled (new config)
        #check if we find "config" disabled for excel
        self.assertTrue('config' in self.config.get_disabled_app_list('addin', 'excel'))

        #enable config for excel, config must not be in disabled list
        self.config.enable_app('addin', 'config', 'excel')
        self.assertFalse('config' in self.config.get_disabled_app_list('addin', 'excel'))

    def test_get_enabled_app_list(self):
        #app_type tests doesn't exist, must return None
        self.assertIsNone(self.config.get_enabled_app_list("tests", "excel"))

        #all apps are disabled (new config)
        #config must not be in disabled list
        self.assertFalse('config' in self.config.get_enabled_app_list('addin', 'excel'))

        #check if we find "config" enabled for excel
        self.config.enable_app('addin', 'config', 'excel')
        self.assertTrue('config' in self.config.get_enabled_app_list('addin', 'excel'))

    def test_print_app_list(self):
        self.config.print_app_list('test')
        self.config.print_app_list('addin')
        self.config.print_app_list('docapp')
        self.assertTrue(True)

    def test_reset(self):
        #reset is already called in setUp
        #config db must not exist
        import os
        self.assertFalse(os.path.exists(os.getenv("USERPROFILE") + "\\.transition.s3db"))

    def test_update_inventory(self):
        import shutil
        self.evt.last_fired_evt = ''
        self.evt.last_fired_evt_args = dict()

        #remove Manifest file in config app directory
        if os.path.exists(self.config.get_app_info('addin', 'config')['path'] + "\\Manifest"):
            os.remove(self.config.get_app_info('addin', 'config')['path'] + "\\Manifest")

        #populate database. must not fire events
        self.config.update_inventory(fire_event=False)
        self.assertEqual('', self.evt.last_fired_evt)
        self.assertEqual(0, len(self.evt.last_fired_evt_args))

        #all apps must have "unchanged" status
        for item in self.config.get_app_list():
            print('test state of', item['app_type'], item['app_name'])
            self.assertEqual(0, self.config.get_app_state(item['app_type'], item['app_name']))

        #make a fake app
        from transitioncore.configsql import SQL_INSERT_APP_TYPE
        import tests
        os.mkdir(self.test_dir)
        with open(self.test_dir + "\\__init__.py", "w") as f:
            f.write("""
class FakeApp():
    def run():
        pass

app_class = FakeApp
com_app = ('excel', 'unknown com app')
app_type = 'tests'
""")

        #insert "tests" app_type
        sqlitecnx = self.config._sqlite
        cursor = sqlitecnx.cursor()
        cursor.execute(SQL_INSERT_APP_TYPE, ("tests", tests.__path__[0]))
        sqlitecnx.commit()

        #Now an update must fire on_app_add
        self.config.update_inventory(app_type='tests')
        self.assertEqual('on_app_add', self.evt.last_fired_evt)
        self.assertEqual('tests', self.evt.last_fired_evt_args['app_type'])
        self.assertEqual('test_files', self.evt.last_fired_evt_args['app_name'])

        #modify app, update must fire on_app_update
        with open(self.test_dir + "\\__init__.py", "w") as f:
            f.write("""
class FakeApp():
    def run():
        print("hello world !")

app_class = FakeApp
com_app = ('excel', 'unknown com app')
app_type = 'tests'
""")
        self.config.update_inventory(app_type='tests')
        self.assertEqual('on_app_update', self.evt.last_fired_evt)
        self.assertEqual('tests', self.evt.last_fired_evt_args['app_type'])
        self.assertEqual('test_files', self.evt.last_fired_evt_args['app_name'])

        #enable config for excel
        self.config.enable_app('addin', 'config', 'excel')

        #sign fake app, and copy Manifest file into config dir to corrupt it
        self.config.write_app_manifest(self.test_dir)

        shutil.move(self.test_dir + "\\Manifest", self.config.get_app_info('addin', 'config')['path'])

        #an update must disable corrupted config app and fire on_app_disable
        self.assertEqual(-1, self.config.get_app_state('addin', 'config'))
        self.config.update_inventory('addin')

        self.assertEqual('addin', self.evt.last_fired_evt_args['app_type'])
        self.assertEqual('config', self.evt.last_fired_evt_args['app_name'])
        self.assertEqual('on_app_disable', self.evt.last_fired_evt)
        self.assertEqual(list(('excel', )), list(self.evt.last_fired_evt_args['com_app_tuple']))
        os.remove(self.config.get_app_info('addin', 'config')['path'] + "\\Manifest")

        #remove fake app, update must fire on_app_del
        shutil.rmtree(self.test_dir)
        self.config.update_inventory(app_type='tests')
        self.assertEqual('on_app_del', self.evt.last_fired_evt)
        self.assertEqual('tests', self.evt.last_fired_evt_args['app_type'])
        self.assertEqual('test_files', self.evt.last_fired_evt_args['app_name'])

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

        expected_digests = self.config.write_app_manifest(self.test_dir)
        self.assertIsInstance(expected_digests, dict)

        self.assertDictEqual(expected_digests, self.config._generate_app_digests(self.test_dir))

if __name__ == '__main__':
    unittest.main()