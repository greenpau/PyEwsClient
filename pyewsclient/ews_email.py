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
import datetime;
import traceback;
from lxml import etree;
import pprint;
import re;

class EWSEmail:
    '''Represents Microsoft Office 365 EWS Email Draft Object.'''


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
                if p == 'error' and self.log[x][level] not in ['CRIT', 'ERROR']:
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
            ''' Display Email Fields '''
            for i in self.skel:
                if isinstance(self.skel[i], str):
                    fld = self.skel[i];
                elif isinstance(self.skel[i], list):
                    fld = ', '.join(self.skel[i]);
                elif isinstance(self.skel[i], dict):
                    fld = str(self.skel[i]);
                else:
                    fld = str(type(self.skel[i]));
                print("{0:26s} | {1:s} | {2:s} | {3:s}".format( str(datetime.datetime.now()), 
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
        ''' Create SOAP Request Body for Email '''

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


        DRAFT = etree.Element(NS_SOAP_ENV + "Envelope", nsmap={'t': NS_EWS_TYPES_URI, 'soap': NS_SOAP_ENV_URI});
        DRAFT_B = etree.SubElement(DRAFT, NS_SOAP_ENV + "Body");
        DRAFT_B_CI = etree.SubElement(DRAFT_B, 'CreateItem');
        DRAFT_B_CI.attrib['xmlns'] = NS_EWS_MESSAGES_URI;

        DRAFT_B_CI.attrib['MessageDisposition'] = 'SaveOnly';
        DRAFT_B_CI_SF = etree.SubElement(DRAFT_B_CI, 'SavedItemFolderId');
        DRAFT_B_CI_DF = etree.SubElement(DRAFT_B_CI_SF, NS_EWS_TYPES + 'DistinguishedFolderId');
        DRAFT_B_CI_DF.attrib['Id'] = 'drafts';

        DRAFT_B_CI_IT = etree.SubElement(DRAFT_B_CI, 'Items');
        DRAFT_B_CI_IT_MSG = etree.SubElement(DRAFT_B_CI_IT, NS_EWS_TYPES + 'Message');

        # Reference for ItemClass: http://msdn.microsoft.com/en-us/library/office/ff861573.aspx
        DRAFT_B_CI_IT_MSG_IC = etree.SubElement(DRAFT_B_CI_IT_MSG, NS_EWS_TYPES + 'ItemClass');
        DRAFT_B_CI_IT_MSG_IC.text = 'IPM.Note';

        if 'subject' in self.skel:
            DRAFT_B_CI_IT_MSG_SB = etree.SubElement(DRAFT_B_CI_IT_MSG, NS_EWS_TYPES + 'Subject');
            DRAFT_B_CI_IT_MSG_SB.text = self.skel['subject'];

        if 'sensitivity' in self.skel:
            DRAFT_B_CI_IT_MSG_SENS = etree.SubElement(DRAFT_B_CI_IT_MSG, NS_EWS_TYPES + 'Sensitivity');
            DRAFT_B_CI_IT_MSG_SENS.text = self.skel['sensitivity'];

        if 'body' in self.skel:
            DRAFT_B_CI_IT_MSG_BD = etree.SubElement(DRAFT_B_CI_IT_MSG, NS_EWS_TYPES + 'Body');
            DRAFT_B_CI_IT_MSG_BD.attrib['BodyType'] = 'Text';
            DRAFT_B_CI_IT_MSG_BD.text = self.skel['body'];            

        if 'importance' in self.skel:
            DRAFT_B_CI_IT_MSG_IMPT = etree.SubElement(DRAFT_B_CI_IT_MSG, NS_EWS_TYPES + 'Importance');
            DRAFT_B_CI_IT_MSG_IMPT.text = self.skel['importance'];

        for i in [('recipients', 'ToRecipients'), ('cc', 'CcRecipients'), ('bcc', 'BccRecipients')]:
            if i[0] not in self.skel:
                continue;
            vars()['DRAFT_B_CI_IT_MSG_' + i[0]] = etree.SubElement(DRAFT_B_CI_IT_MSG, NS_EWS_TYPES + i[1]);
            if isinstance(self.skel[i[0]], (list)) and len(self.skel[i[0]]) > 0:
                for j in self.skel[i[0]]:
                    vars()['DRAFT_B_CI_IT_MSG_' + i[0] + '_MB_' + str(self.skel[i[0]].index(j))] = etree.SubElement(vars()['DRAFT_B_CI_IT_MSG_' + i[0]], NS_EWS_TYPES + 'Mailbox');                      
                    vars()['DRAFT_B_CI_IT_MSG_' + i[0] + '_MB_' + str(self.skel[i[0]].index(j)) + '_EXTRA'] = etree.SubElement(vars()['DRAFT_B_CI_IT_MSG_' + i[0] + '_MB_' + 
                                                                                                 str(self.skel[i[0]].index(j))], NS_EWS_TYPES + 'EmailAddress');
                    vars()['DRAFT_B_CI_IT_MSG_' + i[0] + '_MB_' + str(self.skel[i[0]].index(j)) + '_EXTRA'].text = j;

        if 'read_receipt' in self.skel:
            if self.skel['read_receipt'] == 'Yes':
                DRAFT_B_CI_IT_MSG_RR = etree.SubElement(DRAFT_B_CI_IT_MSG, NS_EWS_TYPES + 'IsReadReceiptRequested');
                DRAFT_B_CI_IT_MSG_RR.text = 'true';

        if 'delivery_receipt' in self.skel:
            if self.skel['delivery_receipt'] == 'Yes':
                DRAFT_B_CI_IT_MSG_DR = etree.SubElement(DRAFT_B_CI_IT_MSG, NS_EWS_TYPES + 'IsDeliveryReceiptRequested');
                DRAFT_B_CI_IT_MSG_DR.text = 'true';

        if 'sender' in self.skel:
            DRAFT_B_CI_IT_MSG_FROM = etree.SubElement(DRAFT_B_CI_IT_MSG, NS_EWS_TYPES + 'From');
            DRAFT_B_CI_IT_MSG_FROM_MB = etree.SubElement(DRAFT_B_CI_IT_MSG_FROM, NS_EWS_TYPES + 'Mailbox');
            DRAFT_B_CI_IT_MSG_FROM_MB_EXTRA = etree.SubElement(DRAFT_B_CI_IT_MSG_FROM_MB, NS_EWS_TYPES + 'EmailAddress');
            DRAFT_B_CI_IT_MSG_FROM_MB_EXTRA.text = self.skel['sender'];            

        if 'mark_read' in self.skel:
            DRAFT_B_CI_IT_MSG_RD = etree.SubElement(DRAFT_B_CI_IT_MSG, NS_EWS_TYPES + 'IsRead');
            if self.skel['mark_read'] == 'Yes':
                DRAFT_B_CI_IT_MSG_RD.text = 'true';
            else:
                DRAFT_B_CI_IT_MSG_RD.text = 'false';

        xmlb = b'<?xml version="1.0" encoding="utf-8"?>\n' + etree.tostring(DRAFT, pretty_print=True);
        self.xml =  xmlb.decode("utf-8");
        return;


    def _ews_schema_checks(self, xmlreq):
        ''' XML Schema Validation '''

        if not isinstance(xmlreq, bytes):
            xmlreq = bytes(xmlreq, 'utf-8');
        
        try:
            msg_schema_xsd = os.path.join('/'.join(os.path.abspath(__file__).split('/')[:-1]), 'xml/messages.xsd');
            msg_schema = etree.XMLSchema(file=msg_schema_xsd);
        except Exception as err:
            self._log(str(err), 'CRIT');
            self._log(str(traceback.format_exc()), 'CRIT');
            return 1;

        try:
            xmlreq_valid = msg_schema.validate(etree.fromstring(xmlreq));
        except Exception as err:
            self._log(str(err), 'CRIT');
            self._log(str(traceback.format_exc()), 'CRIT');

        try:
            msg_schema.assertValid(etree.fromstring(xmlreq));
        except Exception as err:
            if self.verbose >= 3:
                self._log(str(err), 'WARN');
                self._log(str(traceback.format_exc()), 'WARN');

        if xmlreq_valid is True:
            if self.verbose >= 4:
                self._log('SOAP Request is valid', 'INFO');
            return 0;
        else:
            if self.verbose >= 3:
                self._log('SOAP Request is invalid', 'WARN');
            return 1;


    def formatting(self, i):
        ''' Defines Email Fortmattting, e.g. plain or html '''

        if not isinstance(i, str):
            self._log('Email formatting property is invalid', 'CRIT');
            return;
       
        if i not in ['plain', 'html']:
            self._log('Email formatting property is neither plain nor html', 'WARN');
            return;

        self.skel['formatting'] = i;
        if self.verbose >= 4:
            self._log('Email formatting property is ' + i, 'INFO');
        
        return;


    def sender(self, i):
        ''' Defines Email Sender or From Field '''

        if not isinstance(i, str):
            self._log('Email sender field is invalid', 'CRIT');
            return;
       
        self.skel['sender'] = i;
        if self.verbose >= 4:
            self._log('Email sender field is ' + i, 'INFO');

        return;

    def recipients(self, i):
        ''' Defines Email Recipients Field '''

        if not isinstance(i, list):
            self._log('Email recipients field is invalid', 'CRIT');
            return;
        
        self.skel['recipients'] = [];

        for e in i:
            m = re.search(r'([\w.]+)@([\w.]+)', e)
            if m:
                self.skel['recipients'].append(m.group(1) + '@' + m.group(2));

        if self.verbose >= 4:
            self._log('Email recipients field is ' + ', '.join(i), 'INFO');

        return;

    def subject(self, i):
        ''' Defines Email Subject Field '''

        if not isinstance(i, str):
            self._log('Email subject field is invalid', 'CRIT');
            return;
       
        self.skel['subject'] = i;
        if self.verbose >= 4:
            self._log('Email subject field is ' + i, 'INFO');

        return;


    def body(self, i):
        ''' Defines Email Body Field '''

        if not isinstance(i, str):
            self._log('Email body field is invalid', 'CRIT');
            return;
       
        self.skel['body'] = i;
        if self.verbose >= 4:
            self._log('Email body field is ' + i, 'INFO');

        return;
            

    def cc(self, i):
        ''' Defines Email Cc  Field '''
        
        if not isinstance(i, list):
            self._log('Email Cc field is invalid', 'CRIT');
            return;
            
        self.skel['cc'] = i;
        if self.verbose >= 4:
            self._log('Email Cc field is ' + ', '.join(i), 'INFO');
            
        return;
            

    def bcc(self, i):
        ''' Defines Email Bcc  Field '''
        
        if not isinstance(i, list):
            self._log('Email Bcc field is invalid', 'CRIT');
            return;
            
        self.skel['bcc'] = i;
        if self.verbose >= 4:
            self._log('Email Bcc field is ' + ', '.join(i), 'INFO');
            
        return;


    def sensitivity(self, i):
        ''' Defines Email sensitivity property, e.g. Normal, Personal, Private, or Confidential '''

        if not isinstance(i, str):
            self._log('Email sensitivity property is invalid', 'CRIT');
            return;
        
        if i not in ['Normal', 'Personal', 'Private', 'Confidential']:
            self._log('Email sensitivity property is not Normal, Personal, Private, or Confidential', 'WARN');
            return;
    
        self.skel['sensitivity'] = i;
        if self.verbose >= 4:
            self._log('Email sensitivity property is ' + i, 'INFO');
            
        return;


    def importance(self, i):
        ''' Defines Email importance property, e.g. High, Normal, or Low '''

        if not isinstance(i, str):
            self._log('Email importance property is invalid', 'CRIT');
            return;

        if i not in ['High', 'Normal', 'Low']:
            self._log('Email importance property is not High, Normal, or Low', 'WARN');
            return;

        self.skel['importance'] = i;
        if self.verbose >= 4:
            self._log('Email importance property is ' + i, 'INFO');

        return;
        

    def delivery_receipt(self, i):
        ''' Defines Email Delivery Receipt Property '''
            
        if not isinstance(i, str):
            self._log('Email Delivery Receipt Property is invalid', 'CRIT');
            return;

        if i not in ['Yes', 'No']:
            self._log('Email Delivery Receipt Property is neither Yes nor No', 'WARN');
            return;
        
        if i == 'Yes':
            self.skel['delivery_receipt'] = i;
            if self.verbose >= 4:
                self._log('Email Delivery Receipt Property is ' + i, 'INFO');

        return;


    def read_receipt(self, i):
        ''' Defines Email Read Receipt Requested Property '''
            
        if not isinstance(i, str):
            self._log('Email Read Receipt Requested Property is invalid', 'CRIT');
            return;

        if i not in ['Yes', 'No']:
            self._log('Email Read Receipt Requested Property is neither Yes nor No', 'WARN');
            return;

        if i == 'Yes':
            self.skel['read_receipt'] = i;
            if self.verbose >= 4:
                self._log('Email Read Receipt Requested Property is ' + i, 'INFO');

        return;


    def mark_read(self, i):
        ''' Defines Email Mark Read Property, e.g. Yes or No '''

        if not isinstance(i, str):
            self._log('Email formatting property is invalid', 'CRIT');
            return;
       
        if i not in ['Yes', 'No']:
            self._log('Email Mark Read Property is neither Yes nor No', 'WARN');
            return;

        self.skel['mark_read'] = i;
        if self.verbose >= 4:
            self._log('Email Mark Read Property is ' + i, 'INFO');
        
        return;


    def __init__(self, verbose=None):
        ''' Initialize Microsoft Office 365 EWS Email Object '''

        self.verbose = verbose;

        self.log = {};
        self._log_id = 0;
        self.error = False;

        self.xml = None;

        self.skel = {};

        if self.verbose > 0:
            self._log( 'Log Level : ' + str(self.verbose), 'INFO');

        return;

