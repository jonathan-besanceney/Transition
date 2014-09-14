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

# SQL

"""
app table. Contains apps info.

app signatures are inspired from gentoo portage ebuilds
- see Lib/hashlib for SHA 256 / 512
- see whirlpool on pypi
"""
SQL_CREATE_APP = """
CREATE TABLE app
(
    name TEXT,
    author TEXT,
    version TEXT,
    description TEXT,
    active BOOL DEFAULT 0,
    path TEXT,
    id_app_type INT,
    SHA256 VARCHAR(64),
    SHA512 VARCHAR(128),
    WHIRLPOOL VARCHAR(128)
)
"""

SQL_INSERT_APP = """
INSERT INTO app VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
"""

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
"""

SQL_DELETE_APP_BY_ID = """
DELETE FROM app WHERE rowid = ?
"""

SQL_SELECT_APP_LIST_BY_TYPE = """
SELECT app.name, author, version, description, active, app.path, SHA256, SHA512, WHIRLPOOL
FROM app
    INNER JOIN app_type ON app.id_app_type = app_type.rowid
WHERE app_type.name = ?
AND app.active IN (?)
"""

SQL_SELECT_APP_BY_PATH = """
SELECT app.rowid, app.name, author, version, description, active, app.path, SHA256, SHA512, WHIRLPOOL
FROM app
WHERE app.path = ?
"""

SQL_SELECT_APP = """
SELECT app.rowid, app.name, author, version, description, active, app.path, SHA256, SHA512, WHIRLPOOL
FROM app
 INNER JOIN app_type ON app.id_app_type = app_type.rowid
WHERE app_type.name = ?
 AND app.name = ?
"""

SQL_SELECT_APP_ID = """
SELECT app.rowid
FROM app, app_type
WHERE app.id_app_type = app_type.rowid
AND app_type.name = ?
AND app.name = ?
"""

SQL_SELECT_APP_TYPE_ID = """
SELECT app_type.rowid
FROM app_type
WHERE app_type.name = ?
"""

SQL_CREATE_APP_WORKS_WITH_COM_APP = """
CREATE TABLE app_works_with_com_app
(
    id_app INT,
    id_com_app INT
)
"""

SQL_INSERT_APP_WORKS_WITH_COM_APP = """
INSERT INTO app_works_with_com_app VALUES (?, ?)
"""

SQL_DELETE_APP_WORKS_WITH_COM_APP_BY_ID = """
DELETE FROM app_works_with_com_app WHERE id_app = ?
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
app_type table
app_type.name :
- documentapp, launched with particular docs
- complugin, launched with com application
app_type.path refers to package
"""
SQL_CREATE_APP_TYPE = """
CREATE TABLE app_type
(
    name TEXT NOT NULL UNIQUE,
    path TEXT NOT NULL UNIQUE
)
"""

SQL_INSERT_APP_TYPE = """
INSERT INTO app_type VALUES (?, ?)
"""

SQL_SELECT_APP_TYPE_PATH = """
SELECT path FROM app_type WHERE name = ?
"""

SQL_INSERT_COM_APP = """
INSERT INTO com_app VALUES (?)
"""

SQL_SELECT_COM_APP = """
SELECT rowid, short_name
FROM com_app
WHERE short_name = ?
"""