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
import io;
import datetime;
import traceback;
from lxml import etree;
import pprint;
import base64;
from random import randint;
from pyewsclient import EWSXmlSchemaValidator;


class EWSAttachment:
    '''Represents Microsoft Office 365 EWS Email Attachment Object.'''

    def _exit(self, lvl=0):
        if self.log:
            self.show('log', 'error');
        if lvl == 1:
            exit(1);
        else:
            exit(0);


    def _log(self, msg='TEST', lvl='INFO'):
        ''' Logging '''
        lvls={'DEBUG': 5, 'CRIT': 4, 'ERROR': 3, 'WARN': 2, 'INFO': 1};
        cls = str(type(self).__name__);
        func = str(sys._getframe(1).f_code.co_name);
        ts = str(datetime.datetime.now());
        for xmsg in msg.split('\n'):
            if self.error is not True and lvls[lvl] in [3, 4]:
                self.error = True;
            self._log_id += 1;
            self.log[self._log_id] = {'ts': ts, 'function': __file__.split('/')[-1] + '->' + cls + '.' + func + '()', 'level': lvl, 'text': xmsg}
        return;


    def show(self, t=None, p=None):
        ''' Display information '''
        if t == 'log':
            ''' Display log buffer '''
            for x in self.log:
                if p == 'error' and self.log[x]['level'] not in ['CRIT', 'ERROR']:
                    continue;
                print("{0:26s} | {1:s} | {2:s} | {3:s}".format(self.log[x]['ts'],
                                                               self.log[x]['function'],
                                                               self.log[x]['level'],
                                                               self.log[x]['text']));
        elif t == 'request':
            ''' Display SOAP XML Request '''
            if self.xml:
                for x in self.xml.split('\n'):
                    print("{0:26s} | {1:s} | {2:s} | {3:s}".format(str(datetime.datetime.now()),
                        __file__.split('/')[-1] + '->' + str(type(self).__name__) + '.' + str(sys._getframe(1).f_code.co_name) + '()',
                        'INFO', x));
        elif t == 'fields':
            ''' Display Attachment Fields '''
            for i in self.skel:
                if isinstance(self.skel[i], str):
                    fld = self.skel[i];
                elif isinstance(self.skel[i], list):
                    fld = ', '.join(self.skel[i]);
                elif isinstance(self.skel[i], dict):                                                 
                    fld = str(self.skel[i]);
                else:
                    fld = str(type(self.skel[i]));
                print("{0:26s} | {1:s} | {2:s} | {3:s}".format(str(datetime.datetime.now()),
                    __file__.split('/')[-1] + '->' + str(type(self).__name__) + '.' + str(sys._getframe(1).f_code.co_name) + '()',
                    'INFO', i + ' => ' + fld));
        else:
            pass;
        return;


    def clear(self, t=None, p=None):
        if t == 'log':
            self.log.clear();
        return;


    def finalize(self):
        ''' Create SOAP Request Body for Email Attachment'''

        if len(self.skel['attachments']) < 1:
            self._log('No attachments', 'CRIT');
            return;

        NS_SOAP_ENV = "{http://schemas.xmlsoap.org/soap/envelope/}";
        NS_SOAP_ENV_URI = "http://schemas.xmlsoap.org/soap/envelope/";
        NS_XSI = "{http://www.w3.org/1999/XMLSchema-instance}";
        NS_XSI_URI = "http://www.w3.org/1999/XMLSchema-instance";
        NS_XSD = "{http://www.w3.org/1999/XMLSchema}";
        NS_XSD_URI = "http://www.w3.org/1999/XMLSchema";
        NS_WSA = "{http://www.w3.org/2005/08/addressing}";
        NS_WSA_URI = "http://www.w3.org/2005/08/addressing";
        NS_EWS_AUTO = "{http://schemas.microsoft.com/exchange/2010/Autodiscover}";
        NS_EWS_AUTO_URI = "http://schemas.microsoft.com/exchange/2010/Autodiscover";
        NS_EWS_TYPES = "{http://schemas.microsoft.com/exchange/services/2006/types}";
        NS_EWS_TYPES_URI = 'http://schemas.microsoft.com/exchange/services/2006/types';
        NS_EWS_MESSAGES = "{http://schemas.microsoft.com/exchange/services/2006/messages}";
        NS_EWS_MESSAGES_URI = "http://schemas.microsoft.com/exchange/services/2006/messages";

        NSM = {'xsi': NS_XSI_URI, 
               'xsd': NS_XSD_URI, 
               't': NS_EWS_TYPES_URI, 
               'soap': NS_SOAP_ENV_URI};

        ATTACH = etree.Element(NS_SOAP_ENV + "Envelope", nsmap=NSM);
        ATTACH_B = etree.SubElement(ATTACH, NS_SOAP_ENV + "Body");
        ATTACH_B_CA = etree.SubElement(ATTACH_B, 'CreateAttachment');

        ATTACH_B_CA.attrib['xmlns'] = NS_EWS_MESSAGES_URI;
        #ATTACH_B_CA.attrib['xmlns:t'] = NS_EWS_TYPES_URI;

        ATTACH_B_CA_PID = etree.SubElement(ATTACH_B_CA, 'ParentItemId');
        ATTACH_B_CA_PID.attrib['Id'] = self.id;
        ATTACH_B_CA_PID.attrib['ChangeKey'] = self.changekey;

        ATTACH_B_CA_AT = etree.SubElement(ATTACH_B_CA, 'Attachments');

        k = 0;
        for i in self.skel['attachments']:
            k += 1;
            vars()['ATTACH_B_CA_AT_FL' + str(k)] = etree.SubElement(ATTACH_B_CA_AT, NS_EWS_TYPES + 'FileAttachment');
            if self.skel['attachments'][i]['Name'] is not None:
                vars()['ATTACH_B_CA_AT_' + str(k) + '_NAME'] = etree.SubElement(vars()['ATTACH_B_CA_AT_FL' + str(k)], NS_EWS_TYPES + 'Name');
                vars()['ATTACH_B_CA_AT_' + str(k) + '_NAME'].text = str(self.skel['attachments'][i]['Name']);
            if self.skel['attachments'][i]['Content'] is not None:
                vars()['ATTACH_B_CA_AT_' + str(k) + '_CONTENT'] = etree.SubElement(vars()['ATTACH_B_CA_AT_FL' + str(k)], NS_EWS_TYPES + 'Content');
                vars()['ATTACH_B_CA_AT_' + str(k) + '_CONTENT'].text = self.skel['attachments'][i]['Content'];

        xmlb = b'<?xml version="1.0" encoding="utf-8"?>\n' + etree.tostring(ATTACH, pretty_print=True);
        self.xml = xmlb.decode("utf-8");
        return;
       

    def add(self, fp=None, fn=None):
        ''' Defines Email Attachments '''

        if fp is None:
            self._log('Email attachment file path field is invalid, because it is None', 'ERROR');
            return;

        if not isinstance(fn, str) and fn is not None:
            self._log('Email attachment file name field is invalid, because it is not string or None', 'ERROR');
            return;

        if not isinstance(fp, str) and not isinstance(fp, bytes) and not isinstance(fp, io.IOBase):
            self._log('Email attachment file path field is invalid, because it is not string or bytes or file handle', 'ERROR');
            return;

        if isinstance(fp, str):
            if not os.path.isfile(fp):
                self._log('Email attachment ' + fp + ' either does not exist or is not a file', 'ERROR');
                return;
            if not os.access(fp, os.R_OK):
                self._log('Email attachment ' + fp + ' is not readable', 'ERROR');
                return;
            if fn is None:
                fn = os.path.basename(fp)
            with open(fp, "rb") as f:
                fc = base64.b64encode(f.read());
        elif isinstance(fp, io.IOBase):
            fn = os.path.basename(fp.name);
            fc = base64.b64encode(bytes(fp.read(), 'utf-8'));
        else:
            if fn is None:
                fn = 'noname.' + randint(10000,20000);
            fc = fp;

        try:
            f = self.skel['attachments'][self.aid];
        except:
            self.skel['attachments'][self.aid] = {};

        try:
            self.skel['attachments'][self.aid]['Name'] = fn
            self.skel['attachments'][self.aid]['Content'] = fc
            if self.verbose >= 5:
                self._log('email attachment ' + str(fn) + ' was added successfully', 'INFO');
        except:
            self._log('failed to add email attachment ' + str(fn), 'ERROR');
        self.aid += 1;
        return;


    def __init__(self, id=None, changekey=None, verbose=None):
        ''' Initialize Microsoft Office 365 EWS Email Attachment Object '''

        self.verbose = verbose;

        self.log = {};
        self._log_id = 0;
        self.error = False;

        self.xml = None;

        self.aid = 0;
        self.skel = {};
        self.skel['attachments'] = {};

        if self.verbose > 0:
            self._log( 'Log Level : ' + str(self.verbose), 'INFO');

        if id is not None:
            self.id = id

        if changekey is not None:
            self.changekey = changekey

        return;
