@echo off

if exist .\ACLibImportWizard.accdb (
set /p CopyFile=ACLibImportWizard.accdb exists .. overwrite with access-add-in\ACLibImportWizard.accda? [Y/N]:
) else (
set CopyFile=Y
)

if /I %CopyFile% == Y (
	echo File is copied ...
) else (
	echo Batch is cancelled
	pause
	exit
)

copy .\access-add-in\ACLibImportWizard.accda ACLibImportWizard.accdb

timeout 2