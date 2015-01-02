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

from setuptools import setup
from codecs import open
from os import path

pkg_dir = path.abspath(path.dirname(__file__));
pkg_name = 'PyEwsClient';
pkg_ver = '1.0';
pkg_summary = 'Microsoft Office 365 EWS (Exchange Web Services) Client Library';
pkg_url = 'https://github.com/greenpau/' + pkg_name;
pkg_download_url = 'http://pypi.python.org/packages/source/' + pkg_name[0] + '/' + pkg_name + '/' + pkg_name + '-' + pkg_ver + '.tar.gz';
pkg_author = 'Paul Greenberg';
pkg_author_email = 'paul@greenberg.pro';
pkg_packages = [pkg_name.lower()];
pkg_requires = ['lxml', 'requests'];

with open(path.join(pkg_dir, 'README.rst'), encoding='utf-8') as f:
    pkg_long_description = f.read();

setup(
    name=pkg_name,
    version=pkg_ver,
    description=pkg_summary,
    long_description=pkg_long_description,
    url=pkg_url,
    download_url=pkg_download_url,
    author=pkg_author,
    author_email=pkg_author_email,
    license='GPLv3',
    platforms='any',
    classifiers=[
        'Development Status :: 2 - Pre-Alpha',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.0',
        'Programming Language :: Python :: 3.1',
        'Programming Language :: Python :: 3.2',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
        'Operating System :: OS Independent',
        'Topic :: Communications :: Email',
        'Topic :: Communications :: Email :: Address Book',
        'Topic :: Communications :: Email :: Email Clients (MUA)',
        'Topic :: Software Development :: User Interfaces',
        'Intended Audience :: End Users/Desktop',
        'Intended Audience :: Developers',
        'Intended Audience :: System Administrators',
        'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
        'License :: OSI Approved :: GNU General Public License v3 or later (GPLv3+)',
    ],
    packages=pkg_packages,
    install_requires=pkg_requires,
);
