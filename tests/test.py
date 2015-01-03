#!/usr/bin/env python

#------------------------------------------------------------------------------------------#
# File:      test.py                                                                       #
# Purpose:   PyEwsClient - Microsoft Office 365 Client Library Testing Tool                #
# Author:    Paul Greenberg                                                                #
# Version:   1.0                                                                           #
# Copyright: (c) 2013 Paul Greenberg <paul@greenberg.pro>                                  #
# -----------------------------------------------------------------------------------------#

import os;
import sys;
if sys.version_info[0] < 3:
    sys.stderr.write(os.path.basename(__file__) + ' requires Python 3 or higher.\n');
    sys.stderr.write("python3 " + __file__ + '\n');
    exit(1);
sys.path.append(os.path.join('/'.join(os.path.abspath(__file__).split('/')[:-2])));
import argparse;
import pprint;
import datetime;
import traceback;

try:
    from pyewsclient import EWSSession, EWSEmail, EWSAttachment;
except Exception as err:
    for e in err.args:
        print('%-26s | %s | %s | %s' % (str(datetime.datetime.now()), __file__.split('/')[-1] + '->global()', str(type(err).__name__), str(e)));
    sys.exit(1);

def main():
    func = 'main()';
    descr = 'PyEwsClient - Microsoft Office 365 Client Library Testing Tool\n\n';
    descr += 'examples:\n';
    descr += ' python3 tests/test.py -u email@office365.com -p password --autodiscover -l 5\n';
    descr += ' python3 tests/test.py -u email@office365.com -p password --autodiscover \ \n';
    descr += '                       -a tests/attach1.txt -a tests/attach2.txt -l 1 \n';
    descr += ' python3 tests/test.py --help';
    epil = 'documentation:\n https://github.com/greenpau/PyEwsClient\n\n';
    parser = argparse.ArgumentParser(formatter_class=argparse.RawDescriptionHelpFormatter,description=descr, epilog=epil);
    conn_group = parser.add_argument_group('network connectivity arguments');
    conn_group_sub = conn_group.add_mutually_exclusive_group(required=True);
    conn_group_sub.add_argument('-s', '--server', dest='isrv', metavar='SERVER', type=str, help='Office 365 Server');
    conn_group_sub.add_argument('--autodiscover', dest='iauto', action='store_true', help='Office 365 Autodiscovery On');
    auth_group = parser.add_argument_group('authentication arguments')
    auth_group.add_argument('-u', '--user', dest='iuser', metavar='USERNAME', type=str, required=True, help='Office 365 Username');
    auth_group.add_argument('-p', '--password', dest='ipass', metavar='PASSWORD', type=str, required=True, help='Office 365 Password');
    mail_group = parser.add_argument_group('email arguments');
    mail_group.add_argument('-a', '--attachment', dest='iatt', metavar='ATTACHMENT', action='append', type=argparse.FileType('r'), help='Email Attachment(s)');
    parser.add_argument('-l', '--log-level', dest='ilog', metavar='LOG_LEVEL', type=int, default=0, choices=range(1, 6), help='log level (default: 0)');
    args = parser.parse_args();

    try:
        ''' Step 1: Initialize Office 365 Session '''
        ews = EWSSession(args.iuser, args.ipass, args.isrv, args.ilog);
        if ews.log:
            ews.show('log');
            ews.clear('log');
        if ews.error:
            raise RuntimeError('EWS Connectivity Issues');

        ''' Step 2: Create email draft object '''
        email = EWSEmail(args.ilog);
        email.formatting('plain');
        email.sender(args.iuser);
        email.recipients(['to1@microsoft.com', 'to2@microsoft.com']);
        email.subject('Sample Subject');
        email.body('Sample Body');
        email.cc(['cc1@microsoft.com', 'cc2@microsoft.com', 'cc3@microsoft.com']);
        email.bcc(['bcc1@microsoft.com', 'bcc2@microsoft.com']);
        email.sensitivity('Private');
        email.importance('High');
        email.delivery_receipt('Yes');
        email.read_receipt('Yes');
        email.mark_read('No');
        email.finalize();
        if args.ilog > 0: 
            email.show('request');
            email.show('fields');
        if email.log:
            ews.show('log');
            ews.clear('log');
        if email.error:
            raise RuntimeError('Local Email Drafting Issues');

        ''' Step 3: Submit email draft to Office 365 '''
        ews.submit(email.xml, 'save_only');
        if ews.log:
            ews.show('log');
            ews.clear('log');
        if ews.error:
            raise RuntimeError('EWS Endpoint Email Submission Issues');
        if ews.id is None or ews.changekey is None:
            raise RuntimeError('EWS Endpoint Responded with missing Email Id or ChangeKey');

        ''' Step 4: Add file attachments to the draft '''
        if args.iatt is not None:
            attachment = EWSAttachment(ews.id, ews.changekey, args.ilog);
            for a in args.iatt:
                attachment.add(a);
            attachment.finalize();
            if args.ilog > 0: 
                attachment.show('request');
                attachment.show('fields');
            if attachment.log:
                attachment.show('log');
                attachment.clear('log');
            if attachment.error:
                raise RuntimeError('Local Attachment Processing Issues');

        ''' Step 5: Submit attachment(s) to the email draft '''
        if args.iatt is not None:
            ews.submit(attachment.xml);
            if ews.log:
                ews.show('log');
                ews.clear('log');
            if ews.error:
                raise RuntimeError('EWS Endpoint Attachment Submission Issues');

        print('email was sent successfully');

    except Exception as err:
        for e in err.args:
            print('%-26s | %s | %s | %s' % (str(datetime.datetime.now()), __file__.split('/')[-1] + '->' + func, str(type(err).__name__), str(e)));
        if args.ilog == 5:
            for i in str(traceback.format_exc()).splitlines():
                print('%-26s | %s | %s ' % (str(datetime.datetime.now()), __file__.split('/')[-1] + '->' + func, i));
        return;


if __name__ == '__main__':
    main();
