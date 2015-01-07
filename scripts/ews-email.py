#!/usr/bin/env python

#------------------------------------------------------------------------------------------#
# File:      ews-email.py                                                                  #
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
    descr += ' python3 ' + str(__file__) + ' -u email@office365.com -p password --autodiscover -l 5\n';
    descr += ' python3 ' + str(__file__) + ' -u email@office365.com -p password --autodiscover \ \n';
    descr += '  --to "to1@microsoft.com" --to "to2@microsoft.com" \ \n';
    descr += '  --cc "cc1@microsoft.com" --cc "cc2@microsoft.com" --cc "cc3@microsoft.com" \ \n';
    descr += '  --bcc "bcc1@microsoft.com" --bcc "bcc2@microsoft.com" \ \n';
    descr += '  --subject "Sample Subject" --body "Sample Body" \ \n';
    descr += '  --format plain --sensitivity "Confidential" \ \n';
    descr += '  --importance "High" --delivery-receipt --read-receipt --mark-read \ \n';
    descr += '  --attach scripts/attach1.txt --attach scripts/attach2.txt -l 1 \n';
    descr += ' python3 ' + str(__file__) + ' --help';
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
    mail_group.add_argument('--to', dest='ito', metavar='TO', action='append', type=str, help='Email Receipient(s) To:');
    mail_group.add_argument('--cc', dest='icc', metavar='CC', action='append', type=str, help='Email Receipient(s) Cc:');
    mail_group.add_argument('--bcc', dest='ibcc', metavar='BCC', action='append', type=str, help='Email Receipient(s) Bcc:');
    mail_group.add_argument('--subject', dest='isub', metavar='SUBJECT', type=str, required=True, help='Email Subject');
    mail_group.add_argument('--body', dest='ibdy', metavar='BODY', type=str, required=True, help='Email Body');
    mail_group.add_argument('--attach', dest='iatt', metavar='ATTACHMENT', action='append', type=argparse.FileType('r'), help='Email Attachment(s)');
    mail_group.add_argument('--format', dest='ifmt', metavar='FORMAT', type=str, choices=['plain', 'html'], help='Email Format (plain or html)');
    mail_group.add_argument('--sensitivity', dest='isns', metavar='LEVEL', type=str, 
                            choices=['Normal', 'Personal', 'Private', 'Confidential'], required=False, 
                            help='Email Sensitivity, e.g. Normal, Personal, Private, Confidential');
    mail_group.add_argument('--importance', dest='iimp', metavar='LEVEL', type=str, 
                            choices=['High', 'Normal', 'Low'], required=False,
                            help='Email Importance, e.g. High, Normal, Low');
    mail_group.add_argument('--delivery-receipt', dest='irqd', action='store_true', help='Request Delivery Receipt');
    mail_group.add_argument('--read-receipt', dest='irqr', action='store_true', help='Request Read Receipt');
    mail_group.add_argument('--mark-read', dest='imrd', action='store_true', help='Mark Read');

    parser.add_argument('-l', '--log-level', dest='ilog', metavar='LEVEL', type=int, default=0, choices=range(1, 6), help='log level (default: 0, max: 5)');
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
        if args.ifmt is not None:
            email.formatting(args.ifmt);
        else:
            email.formatting('plain');
        email.sender(args.iuser);
        if args.ito is not None:
            email.recipients(args.ito);
        if args.icc is not None:
            email.cc(args.icc);
        if args.ibcc is not None:
            email.bcc(args.ibcc);
        if args.isub is not None:
            email.subject(args.isub);
        if args.ibdy is not None:
            email.body(args.ibdy);
        if args.isns is not None:
            email.sensitivity(args.isns);
        else:
            email.sensitivity('Normal');
        if args.iimp is not None:
            email.importance(args.iimp);
        else:
            email.importance('Normal');
        if args.irqd:
            email.delivery_receipt('Yes');
        if args.irqr:
            email.read_receipt('Yes');
        if args.imrd:
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
