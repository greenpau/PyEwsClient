#   PyEwsClient - Microsoft Office 365 EWS (Exchange Web Services) Client Library
#   Copyright (C) 2013 Paul Greenberg <paul@greenberg.pro>
#
#   This program is free software: you can redistribute it and/or modify
#   it under the terms of the GNU General Public License as published by
#   the Free Software Foundation, either version 3 of the License, or
#   (at your option) any later version.
#
#   This program is distributed in the hope that it will be useful,
#   but WITHOUT ANY WARRANTY; without even the implied warranty of
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#   GNU General Public License for more details.
#
#   You should have received a copy of the GNU General Public License
#   along with this program.  If not, see <http://www.gnu.org/licenses/>.

import os;
import sys;
import requests;
import datetime;
import traceback;
import requests;
from requests.auth import HTTPBasicAuth;
from lxml import etree;
import pprint;
import re;


class EWSXmlSchemaValidator:
    '''Represents Microsoft Office 365 EWS XML Schema Validation Funstion.'''

    def __init__(self, xmlreq, xmlsch=None):
        ''' XML Schema Validation '''

        self.valid = False;
        self.logs = [];

        if xmlsch is None:
            xmlsch = 'xml/messages.xsd';
        else:
            xmlsch = 'xml/' + xmlsch;

        if not isinstance(xmlreq, bytes):
            xmlreq = bytes(xmlreq, 'utf-8');

        try:
            msg_schema_xsd = os.path.join('/'.join(os.path.abspath(__file__).split('/')[:-1]), xmlsch);
            msg_schema = etree.XMLSchema(file=msg_schema_xsd);
        except Exception as err:
            self.logs.append((str(err), 'ERROR'));
            self.logs.append((str(traceback.format_exc()), 'ERROR'));
            return;

        try:
            xmlreq_valid = msg_schema.validate(etree.fromstring(xmlreq));
            self.valid = True;
        except Exception as err:
            self.logs.append((str(err), 'ERROR'));
            self.logs.append((str(traceback.format_exc()), 'ERROR'));
            self.valid = False;

        try:
            msg_schema.assertValid(etree.fromstring(xmlreq));
            self.valid = True;
        except Exception as err:
            self.logs.append((str(err), 'ERROR'));
            self.logs.append((str(traceback.format_exc()), 'ERROR'));
            self.valid = False;

        if self.valid is not True:
            self.logs.append(('XML document failed XML schema validation', 'ERROR'));
            return;

        self.logs.append(('XML document passed XML schema validation', 'INFO'));
        self.valid = True;
        return;
