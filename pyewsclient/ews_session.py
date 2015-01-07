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
import requests;
from requests.auth import HTTPBasicAuth;
from lxml import etree;
import pprint;
import re;

import base64;
import http.client, urllib.parse;
from urllib.parse import urlparse;

from pyewsclient import EWSXmlSchemaValidator;

#sys.path.append(os.path.join('/'.join(os.path.abspath(__file__).split('/')[:-2])));
#from pyewsclient.ews_helper import EWSHelper;

class EWSSession:
    '''Represents Microsoft Office 365 EWS Session.'''


    def _exit(self, lvl=0):
        if self.log:
            self.show('log', 'error');
        if lvl == 1:
            exit(1);
        else:
            exit(0);


    def _log(self, msg='TEST', lvl='INFO'):
        ''' Logging '''
        if self.verbose < 1 and lvl == 'INFO':
            return;
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
        else:
            pass;
        return;


    def clear(self, t=None, p=None):
        if t == 'log':
            self.log.clear();
        return;


    def _ews_add_cookies(self, c):
        cs = c.split(';');
        if len(cs) > 1:
            m = re.match('([A-Za-z0-9-]+)\=(.*)', cs[0]);
            if m:
                self.cookies[m.group(1)] = m.group(2);
        return;


    def _ews_inject_cookies(self, h):
        if len(self.cookies) > 0:
            cs = [];
            for c in self.cookies:
                cs.append(c + '=' + self.cookies[c]);
            h['Cookie'] = '; '.join(cs);
        return h;


    def _ews_urlsplit(self, p, s):
        ''' Parses URL '''
        o = urlparse(s);
        if p == 'prefix':
            return(str(o.scheme));
        if p == 'host':
            return(str(o.hostname));
        if p == 'path':
            return(str(o.path));
        return None;


    def _ews_remove_xml_header(self, s):
        s = s.replace('<?xml version="1.0" encoding="utf-8"?>\r\n', '');
        s = s.replace('<?xml version="1.0" encoding="utf-8"?>', '');
        s = s.replace('<Response xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a">', 
                      '<Response xmlns="http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006">');
        return(s);


    def _ews_autod_request_builder(self):
        ''' SOAP Request Builder '''

        SDIS = etree.Element('Autodiscover');
        SDIS.attrib['xmlns'] = 'http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006';
        SDIS_REQ = etree.SubElement(SDIS, 'Request');
        SDIS_REQ_EMAIL = etree.SubElement(SDIS_REQ, 'EMailAddress');
        SDIS_REQ_EMAIL.text = self.username;
        SDIS_REQ_ARS = etree.SubElement(SDIS_REQ, 'AcceptableResponseSchema');
        SDIS_REQ_ARS.text = 'http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a';
        reqb = b'<?xml version="1.0" encoding="utf-8"?>\n' + etree.tostring(SDIS, pretty_print=True);
        req = reqb.decode("utf-8");
        return req;


    def _ews_send_and_save_request_builder(self):
        ''' SOAP Request Builder '''

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
               'm': NS_EWS_MESSAGES_URI,
               't': NS_EWS_TYPES_URI,
               'soap': NS_SOAP_ENV_URI};

        EMAIL = etree.Element(NS_SOAP_ENV + 'Envelope', nsmap=NSM);
        EMAIL_B = etree.SubElement(EMAIL, NS_SOAP_ENV + 'Body');
        EMAIL_B_SI = etree.SubElement(EMAIL_B, NS_EWS_MESSAGES + 'SendItem');
        EMAIL_B_SI.attrib['SaveItemToFolder'] = 'true';

        EMAIL_B_SI_IDS = etree.SubElement(EMAIL_B_SI, NS_EWS_MESSAGES + 'ItemIds');
        EMAIL_B_SI_ID = etree.SubElement(EMAIL_B_SI_IDS, NS_EWS_TYPES + 'ItemId');
        EMAIL_B_SI_ID.attrib['Id'] = self.id;
        EMAIL_B_SI_ID.attrib['ChangeKey'] = self.changekey;

        EMAIL_B_SI_FI = etree.SubElement(EMAIL_B_SI, NS_EWS_MESSAGES + 'SavedItemFolderId');
        EMAIL_B_SI_DF = etree.SubElement(EMAIL_B_SI_FI, NS_EWS_TYPES + 'DistinguishedFolderId');
        EMAIL_B_SI_DF.attrib['Id'] = 'sentitems';

        reqb = b'<?xml version="1.0" encoding="utf-8"?>\n' + etree.tostring(EMAIL, pretty_print=True);
        req = reqb.decode("utf-8");
        return req;


    def _ews_xml_response_parser(self, stage, url, status, reason, body):
        ''' EWS XML Response Parsing '''
        if stage is None:
            stage = 'non-draft';

        body = self._ews_remove_xml_header(body);
        t = etree.fromstring(body);
        body = etree.tostring(t, pretty_print=True).decode("utf-8");

        if self.verbose >= 4:
            self._log(body, 'INFO');

        NS_EWS_MESSAGES = "{http://schemas.microsoft.com/exchange/services/2006/messages}";
        NS_EWS_MESSAGES_URI = "http://schemas.microsoft.com/exchange/services/2006/messages";
        NS_EWS_TYPES = "{http://schemas.microsoft.com/exchange/services/2006/types}";
        NS_EWS_TYPES_URI = 'ewhttp://schemas.microsoft.com/exchange/services/2006/types';

        NSD = {'m':NS_EWS_MESSAGES_URI, 
               't':NS_EWS_TYPES_URI};
        NSDR = {NS_EWS_MESSAGES:'m',
                   NS_EWS_TYPES:'t'};

        try:
            mResponseMessages = t.xpath('//m:ResponseMessages', namespaces = NSD);
            hc = 0;
            for i in mResponseMessages:
                h = str(hc);
                if self.verbose >= 4:
                    self._log('mResponseMessages[' + h + ']           ' + str(i.tag) + ' / ' + str(i.text))
                for j in list(i):
                    h = str(hc);
                    if self.verbose >= 4:
                        self._log('mResponseMessages[' + h + '] LVL 1 (j) ' + str(j.tag) + ' / ' + str(j.text));
                    for k in list(j):
                        if self.verbose >= 4:
                            self._log('mResponseMessages[' + h + '] LVL 2 (k) -- ' + str(k.tag) + ' / ' + str(k.text));

                        if stage == 'send_and_save' and j.tag == NS_EWS_MESSAGES + 'SendItemResponseMessage':
                            if 'ResponseClass' in j.attrib:
                                if j.attrib['ResponseClass'] == 'Success':
                                    if k.tag == NS_EWS_MESSAGES + 'ResponseCode':
                                        self._log('email was sent successfully. ' + str(k.text));

                        for m in list(k):
                            if self.verbose >= 4:
                                self._log('mResponseMessages[' + h + '] LVL 3 (m) ---- ' + str(m.tag) + ' / ' + str(m.text));
                            for n in list(m):
                                if self.verbose >= 4:
                                    self._log('mResponseMessages[' + h + '] LVL 4 (n) ------ ' + str(n.tag) + ' / ' + str(n.text));
                                
                                if stage == 'save_only' and j.tag == NS_EWS_MESSAGES + 'CreateItemResponseMessage':
                                    if 'ResponseClass' in j.attrib:
                                        if j.attrib['ResponseClass'] == 'Success':
                                            if k.tag == NS_EWS_MESSAGES + 'Items':
                                                if m.tag == NS_EWS_TYPES + 'Message' and n.tag == NS_EWS_TYPES + 'ItemId' and 'ChangeKey' in n.attrib and 'Id' in n.attrib:
                                                    elmN = str(n.tag).replace(NS_EWS_TYPES, NSDR[NS_EWS_TYPES] + ':');
                                                    if self.verbose >= 4:
                                                        self._log('mResponseMessages[' + h + '] LVL 4 (n) ------ ' + elmN + ' has Id (' + n.attrib['Id'] + ') and ChangeKey (' + n.attrib['ChangeKey'] + ') attributes.');
                                                    self.id = str(n.attrib['Id']);
                                                    self.changekey = str(n.attrib['ChangeKey']);

                                if stage == 'attachment' and j.tag == NS_EWS_MESSAGES + 'CreateAttachmentResponseMessage':
                                    if 'ResponseClass' in j.attrib:
                                        if j.attrib['ResponseClass'] == 'Success':
                                            if k.tag == NS_EWS_MESSAGES + 'Attachments':
                                                if m.tag == NS_EWS_TYPES + 'FileAttachment' and n.tag == NS_EWS_TYPES + 'AttachmentId' and 'Id' in n.attrib:
                                                    elmN = str(n.tag).replace(NS_EWS_TYPES, NSDR[NS_EWS_TYPES] + ':');
                                                    if self.verbose >= 4:
                                                        self._log('mResponseMessages[' + h + '] LVL 4 (n) ------ ' + elmN + ' has Id (' + n.attrib['Id'] + ') attribute.');
                                                    if 'RootItemId' in n.attrib:
                                                        if n.attrib['RootItemId'] != self.id:
                                                            self.id = str(n.attrib['RootItemId']);  
                                                    if 'RootItemChangeKey' in n.attrib:
                                                        if n.attrib['RootItemChangeKey'] != self.changekey:
                                                            self.changekey = str(n.attrib['RootItemChangeKey']);  

                    hc += 1;
                hc += 1;
        except Exception as err:
            self._log(str(err), 'ERROR');
            self._log(str(traceback.format_exc()), 'ERROR');
            return;
                                             
        return;


    def _ews_autodiscover(self):
        ''' EWS Autodiscovery 
        URL: autodiscover.outlook.com

        The Autodiscover service may respond with one of two redirection responses: 
         * an HTTP 302 redirect, or
         * a SOAP redirection response. 
        If the response from the Exchange server is an HTTP 302 redirect, 
        the client application should validate that the redirection address 
        is acceptable and then follow the redirection response.

        Microsoft Remote Connectivity Analyzer
          Microsoft Office Outlook Connectivity Tests 
            Outlook Autodiscover 
              https://testconnectivity.microsoft.com/

        [MS-OXDSCLI]: Autodiscover Publishing and Lookup Protocol
          http://download.microsoft.com/download/5/D/D/5DD33FDF-91F5-496D-9884-0A0B0EE698BB/%5BMS-OXDSCLI%5D.pdf

        Importantly, HTTP Location header in HTTP 302 Redirect points to EWS Endpoint!
        '''

        if self.verbose >= 4:
            self._log( 'Autodiscovery On', 'INFO');

        autod_req = self._ews_autod_request_builder();

        exsv = EWSXmlSchemaValidator(autod_req);

        for i in exsv.logs:
            self._log(str(i[0]), str(i[1]));

        if exsv.valid == False:
            self._log('failed ews xml schema validation for autodiscovery', 'ERROR');
            self._exit(1);

        autod_url = 'https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml';
        autod_params = urllib.parse.urlencode({'test1': 123456, '@test2': 'test2', '@test3': 'test3'});
        autod_headers = {'User-Agent': str(self.user_agent),
                         'X-MapiHttpCapability': '1',
                         'Authorization': self.basic_auth,
                         'Content-Type': 'text/xml; charset=utf-8'};
        autod_headers = self._ews_inject_cookies(autod_headers);

        if self.verbose > 4:
            self._log('HTTP REQUEST URL: ' + str(autod_url), 'INFO');

        try:
            if self._ews_urlsplit('prefix', autod_url) == 'https':
                autod_conn = http.client.HTTPSConnection(self._ews_urlsplit('host', autod_url));
            elif self._ews_urlsplit('prefix', autod_url) == 'http':
                autod_conn = http.client.HTTPConnection(self._ews_urlsplit('host', autod_url));
            else:
                return
            if self.verbose >= 5:
                autod_conn.set_debuglevel(self.verbose);
            autod_conn.request("POST", self._ews_urlsplit('path', autod_url), autod_params, autod_headers);
            autod_resp = autod_conn.getresponse();
            autod_resp_headers = autod_resp.getheaders();
            autod_resp_body = self._ews_remove_xml_header(autod_resp.read().decode("utf-8"));
            autod_conn.close();
        except Exception as err:
            self._log(str(err), 'CRIT');
            self._log(str(traceback.format_exc()), 'CRIT');
            return;

        if self.verbose >= 4:
            self._log('HTTP RESPONSE STATUS/REASON: ' + str(autod_resp.status) + '/' + str(autod_resp.reason), 'INFO');

        if isinstance(autod_resp_headers, list):
            for h in autod_resp_headers:
                if self.verbose >= 4:
                    self._log('HTTP RESPONSE HEADER: ' + str(h[0]) + ':   ' + str(h[1]), 'INFO');
        else:
            self._log(autod_url + ' does not respond with headers', 'CRIT');
            return;
 
        if str(autod_resp.status) == '302' and str(autod_resp.reason) == 'Found':
            if autod_resp.getheader('Location') is not None:
                autod_url = autod_resp.getheader('Location');
                if self.verbose > 0:
                    self._log('Candidate EWS Endpoint: ' + self._ews_urlsplit('host', autod_url), 'INFO');
            else:
                self._log(autod_url + ' does not respond with Location header redirect', 'CRIT');
                return;
        else:
            self._log(autod_url + ' does not respond with HTTP 302 Found', 'CRIT');
            return;


        if self.verbose >= 4:
            self._log('HTTP REQUEST URL: ' + str(autod_url), 'INFO');
            self._log('HTTP REQUEST BODY:\n' + str(autod_req), 'INFO');

        try:
            if self._ews_urlsplit('prefix', autod_url) == 'https':
                autod_conn = http.client.HTTPSConnection(self._ews_urlsplit('host', autod_url));
            elif self._ews_urlsplit('prefix', autod_url) == 'http':
                autod_conn = http.client.HTTPConnection(self._ews_urlsplit('host', autod_url));
            else:
                return;
            if self.verbose >= 5:
                autod_conn.set_debuglevel(self.verbose);
            autod_headers = self._ews_inject_cookies(autod_headers);
            autod_conn.request("POST", self._ews_urlsplit('path', autod_url), autod_req, autod_headers);
            autod_resp = autod_conn.getresponse();
            autod_resp_headers = autod_resp.getheaders();
            autod_resp_body = self._ews_remove_xml_header(autod_resp.read().decode("utf-8"));
            autod_conn.close();
        except Exception as err:
            self._log(str(err), 'CRIT');
            self._log(str(traceback.format_exc()), 'CRIT');
            return;

        if self.verbose >= 4:
            self._log('HTTP RESPONSE STATUS/REASON: ' + str(autod_resp.status) + '/' + str(autod_resp.reason), 'INFO');

        if isinstance(autod_resp_headers, list):
            for h in autod_resp_headers:
                if self.verbose >=4 :
                    self._log('HTTP RESPONSE HEADER: ' + str(h[0]) + ':   ' + str(h[1]), 'INFO');
                if str(h[0]) == 'Set-Cookie':
                    self._ews_add_cookies(str(h[1]));
        else:
            self._log(autod_url + ' does not respond with headers', 'CRIT');
            return;

        if isinstance(autod_resp_body, str):
            if len(autod_resp_body) > 20:
                if self.verbose >= 4:
                    self._log('HTTP RESPONSE BODY:\n' + autod_resp_body, 'INFO');
        else:
            self._log(autod_url + ' does not respond with headers', 'CRIT');
            return;

        autod_headers = self._ews_inject_cookies(autod_headers);

        exsv = EWSXmlSchemaValidator(autod_resp_body, 'autodiscover.response.xsd');

        for i in exsv.logs:
            self._log(i[0], i[1]);

        if exsv.valid == True:
            # EwsUrl point to 'https://outlook.office365.com/EWS/Exchange.asmx'
            # However, it should point to 'https://podXXXXX.outlook.com/EWS/Exchange.asmx'
            self.server = 'https://' + self._ews_urlsplit('host', autod_url) + '/EWS/Exchange.asmx';
            return;
        else:
            self.server = None;
            return;


    def submit(self, ews_req=None, ews_stage=None):

        if (ews_stage == 'attachment' or ews_stage == 'send_and_save') and (self.id is None or self.changekey is None):
            self._log('Office 365 email draft token is missing', 'ERROR');
            return;

        if ews_stage == 'send_and_save':
            ews_req = self._ews_send_and_save_request_builder();

        ews_headers = {'User-Agent': str(self.user_agent),
                       'X-MapiHttpCapability': '1',
                       'Authorization': self.basic_auth,
                       'Content-Type': 'text/xml; charset=utf-8'};

        ews_headers = self._ews_inject_cookies(ews_headers);

        if self.verbose >= 4:
            self._log('HTTP REQUEST URL: ' + self.server, 'INFO');
            for h in ews_headers:
                self._log('HTTP REQUEST HEADER: ' + h + ':   ' + ews_headers[h], 'INFO');
            self._log('HTTP REQUEST BODY:\n' + str(ews_req), 'INFO');
                        
        try:
            if self._ews_urlsplit('prefix', self.server) == 'https':
                ews_conn = http.client.HTTPSConnection(self._ews_urlsplit('host', self.server));
            elif self._ews_urlsplit('prefix', self.server) == 'http':
                ews_conn = http.client.HTTPConnection(self._ews_urlsplit('host', self.server));
            else:
                return;
            #if self.verbose >= 5:
            #    ews_conn.set_debuglevel(self.verbose);
            ews_conn.request("POST", self._ews_urlsplit('path', self.server), ews_req, ews_headers);
            ews_resp = ews_conn.getresponse();
            ews_resp_headers = ews_resp.getheaders();
            ews_resp_body = self._ews_remove_xml_header(ews_resp.read().decode("utf-8"));
            ews_conn.close();
        except Exception as err:
            self._log(str(err), 'CRIT');
            self._log(str(traceback.format_exc()), 'CRIT');
            return;

        if self.verbose >= 4:
            self._log('HTTP RESPONSE STATUS/REASON: ' + str(ews_resp.status) + '/' + str(ews_resp.reason), 'INFO');

        if isinstance(ews_resp_headers, list):
            for h in ews_resp_headers:
                if self.verbose >= 4:
                    self._log('HTTP RESPONSE HEADER: ' + str(h[0]) + ':   ' + str(h[1]), 'INFO');
                if str(h[0]) == 'Set-Cookie':
                    self._ews_add_cookies(str(h[1]));
        else:
            self._log(ews_url + ' does not respond with headers', 'CRIT');
            return;

        if isinstance(ews_resp_body, str):
            if len(ews_resp_body) < 20:
                self._log(ews_url + ' text-based output is too short', 'ERROR');
                return;
        else:
            self._log(ews_url + ' does not respond with text-based output', 'ERROR');
            return;

        exsv = EWSXmlSchemaValidator(ews_resp_body);

        for i in exsv.logs:
            self._log(i[0], i[1]);

        if exsv.valid == False:
            self._log('failed ews xml schema validation for ews response', 'ERROR');
            self._exit(1);

        self._ews_xml_response_parser(ews_stage, self.server, str(ews_resp.status), str(ews_resp.reason), ews_resp_body);
        
        return;


    def __init__(self, u=None, p=None, s=None, verbose=0):
        ''' Initialize Microsoft Office 365 Session via SOAP '''

        self.s = requests.Session();
        self.verbose = verbose;

        self.log = {};
        self._log_id = 0;
        self.error = False;

        if isinstance(u, str):
            self.username = u;
        else:
            self._log('expects username parameter to be a string', 'ERROR');
            self._exit(1);

        if isinstance(p, str):
            self.password = p;
        else:
            self._log('expects password parameter to be a string', 'ERROR');
            self._exit(1);

        self.basic_auth = 'Basic ' + base64.urlsafe_b64encode(bytes(self.username + ':' + self.password, 'utf-8')).decode('utf-8');
        self.user_agent = 'Mozilla/5.0 (Windows NT 5.1; rv:31.1) Gecko/20100101 Firefox/31.0';

        self.server = s;
        self.id = None;
        self.changekey = None;
        self.cookies = {};

        if self.verbose > 0:
            self._log( 'Log Level : ' + str(self.verbose), 'INFO');

        if self.server is None:
            self._ews_autodiscover();

        if self.server is None:
            self._log('EWS Endpoint Autodiscovery Failed', 'ERROR');
            self._exit(1);
        
        if self.verbose > 0:
            self._log( 'EWS Endpoint Server: ' + self._ews_urlsplit('host', self.server), 'INFO');


        return;

