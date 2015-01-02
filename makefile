#######################################################################################################
#
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
#
#
#   Prerequisites:
#
#     - Set DEV_BACKUP_DIR environment variable via /etc/profile.d/developer.sh   
#          
#       sudo bash -c "echo '# User Development Environment' > /etc/profile.d/developer.sh"
#       sudo bash -c "echo 'DEV_BACKUP_DIR=~/dev/backup' >> /etc/profile.d/developer.sh"
#       sudo bash -c "echo 'export DEV_BACKUP_DIR' >> /etc/profile.d/developer.sh"
#       sudo bash -c "chmod 644 /etc/profile.d/developer.sh"
#
#     - Set GIT_USER local variable
#
#######################################################################################################

APP_NAME=PyEwsClient
APP_VERSION=1.0
APP_DIR=${APP_NAME}-${APP_VERSION}
DEV_BACKUP_FILE=$(DEV_BACKUP_DIR)/$(APP_NAME).$(APP_VERSION).backup.`date '+%Y%m%d.%H%M%S'`.tar.gz
GIT_USER=$(USER)

all:
	@echo "Running full deployment ..."
	@make build

build:
	@echo "Running package build locally ..."
	@python3 setup.py sdist

pypireg:
	@echo "Running package register ..."
	@python3 setup.py register

pypiupload:
	@echo "Running package build and upload to PyPI ..."
	@python3 setup.py sdist upload

clean:
	@echo "Running cleanup ..."
	@rm -rf ${APP_NAME}.egg-info/ dist/ MANIFEST

backup:
	@echo "Backup ..."
	@mkdir -p ${DEV_BACKUP_DIR}
	@cd ..; tar -cvzf ${DEV_BACKUP_FILE} --exclude='*/.git*' --exclude='*__pycache__*' ${APP_NAME}; cd ${APP_NAME}
	@echo "Completed! Run \"tar -ztvf ${DEV_BACKUP_FILE}\" to verify ..."

git:
	@echo "Running git commit ..."
	@git add -A && git commit -am  "N/A"

github:
	@echo "Running github commit ..."
	@git remote set-url origin git@github.com:${GIT_USER}/${APP_NAME}.git
	@git push origin master
