# ------------------------------------------------------------------------------
# Name:        configsql
# Purpose:     SQL used in configuration.py
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

# ------------------------------------------------------------------------------
#                       SQL CREATE STATEMENTS
# ------------------------------------------------------------------------------

"""
app table. Contains apps info.

app signatures are inspired from gentoo portage ebuilds
- see Lib/hashlib +  pysha3 for SHA 256 / 512
- see whirlpool on pypi
"""
SQL_CREATE_APP = """
CREATE TABLE app
(
    name TEXT,
    author TEXT,
    version TEXT,
    description TEXT,
    path TEXT,
    id_app_type INT,
    SHA256 VARCHAR(64),
    SHA512 VARCHAR(128),
    WHIRLPOOL VARCHAR(128)
)
"""

"""
app_type table
app_type.name :
- docapp, launched with particular docs
- addin, launched with com application
app_type.path refers to package
"""
SQL_CREATE_APP_TYPE = """
CREATE TABLE app_type
(
    name TEXT NOT NULL UNIQUE,
    path TEXT NOT NULL UNIQUE
)
"""

"""
com_app table. Known com application inventory
"""
SQL_CREATE_COM_APP = """
CREATE TABLE com_app
(
    short_name TEXT NOT NULL UNIQUE
)
"""

"""
Link between Transition apps and Office COM Applications.
enabled property is used to make available (or not)
Transition app in considered Office COM Application
"""
SQL_CREATE_APP_WORKS_WITH_COM_APP = """
CREATE TABLE app_works_with_com_app
(
    id_app INT,
    id_com_app INT,
    enabled BOOL DEFAULT 0
)
"""

# ------------------------------------------------------------------------------
#                       SQL INSERT STATEMENTS
# ------------------------------------------------------------------------------

SQL_INSERT_APP = """
INSERT INTO app VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
"""

SQL_INSERT_APP_TYPE = """
INSERT INTO app_type VALUES (?, ?)
"""

SQL_INSERT_COM_APP = """
INSERT INTO com_app VALUES (?)
"""

SQL_INSERT_APP_WORKS_WITH_COM_APP = """
INSERT INTO app_works_with_com_app VALUES (?, ?, ?)
"""

# ------------------------------------------------------------------------------
#                       SQL UPDATE STATEMENTS
# ------------------------------------------------------------------------------
"""
author TEXT,
version TEXT,
description TEXT,
SHA256 VARCHAR(64),
SHA512 VARCHAR(128),
WHIRLPOOL VARCHAR(128)
"""
SQL_UPDATE_APP = """
UPDATE app SET author = ?, version = ?, description = ?, SHA256 = ?, SHA512 = ?, WHIRLPOOL = ?
WHERE app.rowid = ?
"""

SQL_UPDATE_APP_WORKS_WITH_COM_APP = """
UPDATE app_works_with_com_app
SET enabled = ?
WHERE app_works_with_com_app.rowid = (
SELECT app_works_with_com_app.rowid
FROM app_works_with_com_app
    INNER JOIN app ON app_works_with_com_app.id_app =  app.rowid
    INNER JOIN app_type ON app.id_app_type = app_type.rowid
    INNER JOIN com_app ON app_works_with_com_app.id_com_app = com_app.rowid
WHERE  app_type.name = ?
AND app.name = ?
AND com_app.short_name = ?)
"""


# ------------------------------------------------------------------------------
#                       SQL DELETE STATEMENTS
# ------------------------------------------------------------------------------
SQL_DELETE_APP_BY_ID = """
DELETE FROM app WHERE rowid = ?
"""

SQL_DELETE_APP_WORKS_WITH_COM_APP_BY_ID = """
DELETE FROM app_works_with_com_app WHERE id_app = ?
"""

# ------------------------------------------------------------------------------
#                       SQL SELECT STATEMENTS
# ------------------------------------------------------------------------------
SQL_SELECT_APP_LIST_BY_TYPE = """
SELECT app.name, author, version, description, app.path, SHA256, SHA512, WHIRLPOOL
FROM app
    INNER JOIN app_type ON app.id_app_type = app_type.rowid
WHERE app_type.name = ?
"""

SQL_SELECT_APP_LIST_BY_TYPE_AND_COM_APP = """
SELECT app_type.name as app_type, app.name as name, com_app.short_name as com_app, app_works_with_com_app.enabled
FROM app
    INNER JOIN app_type ON app.id_app_type = app_type.rowid
    INNER JOIN app_works_with_com_app ON app.rowid = app_works_with_com_app.id_app
    INNER JOIN com_app ON app_works_with_com_app.id_com_app = com_app.rowid
WHERE app_type.name = ?
AND com_app.short_name IN ({})
"""

SQL_SELECT_APP_LIST_BY_TYPE_AND_COM_APP_AND_STATUS = """
SELECT app_type.name as app_type, app.name as name, com_app.short_name as com_app, app_works_with_com_app.enabled
FROM app
    INNER JOIN app_type ON app.id_app_type = app_type.rowid
    INNER JOIN app_works_with_com_app ON app.rowid = app_works_with_com_app.id_app
    INNER JOIN com_app ON app_works_with_com_app.id_com_app = com_app.rowid
WHERE app_type.name = ?
AND app.name = ?
AND com_app.short_name IN ({})
"""

SQL_SELECT_APP_BY_PATH = """
SELECT app.rowid, app.name, author, version, description, app.path, SHA256, SHA512, WHIRLPOOL,
    app_type.name as app_type, app_type.path as app_type_path
FROM app
    INNER JOIN app_type ON app.id_app_type = app_type.rowid
WHERE app.path = ?
"""

SQL_SELECT_APP = """
SELECT app.rowid, app.name, author, version, description, app.path, SHA256, SHA512, WHIRLPOOL,
    app_type.name as app_type, app_type.path as app_type_path
FROM app
 INNER JOIN app_type ON app.id_app_type = app_type.rowid
WHERE app_type.name = ?
 AND app.name = ?
"""

SQL_SELECT_APP_ID = """
SELECT app.rowid
FROM app
 INNER JOIN app_type ON app.id_app_type = app_type.rowid
WHERE app_type.name = ?
AND app.name = ?
"""

SQL_SELECT_APP_TYPE_ID = """
SELECT app_type.rowid
FROM app_type
WHERE app_type.name = ?
"""

SQL_SELECT_APP_WORKS_WITH_COM_APP = """
SELECT app_type.name as app_type, app.name as app_name, com_app.short_name as com_app, app.description,
    app_works_with_com_app.enabled
FROM app
    INNER JOIN app_type ON app.id_app_type = app_type.rowid
    INNER JOIN app_works_with_com_app ON app.rowid = app_works_with_com_app.id_app
    INNER JOIN com_app ON app_works_with_com_app.id_com_app = com_app.rowid
ORDER BY app_type.name, app.name
"""

SQL_SELECT_APP_TYPE_PATH = """
SELECT path FROM app_type WHERE name = ?
"""

SQL_SELECT_COM_APP = """
SELECT rowid, short_name
FROM com_app
WHERE short_name = ?
"""

SQL_SELECT_ALL_COM_APP = """
SELECT rowid, short_name
FROM com_app
"""