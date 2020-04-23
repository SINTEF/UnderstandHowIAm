==#INFORMATION ABOUT THIS SCRIPT# =====================================================================================
BACKGROUND
This script has been developed as part of the prosject "Forstå hvordan jeg har det" (Understand How I am) 2017-2020.
In the project, heart rate monitors have been used by persons with profound intellectual disability and severe speech,
language or communications difficulties to better understand if and how heart rate can provide more insight for caregivers
into whether the person with severe communication disabilities feels pain, anxiety, stress or other strong emotions as joy.
As part of the project, caregivers have observed behaviors (e.g. a scream or kick) and activities (e.g. meals,
transportation, or physiotherapy sessions) and when they occurred.
PURPOSE
The purpose of this script is to simplify the process of converting files from Garmin heart rate monitor (stored in the
.fit format) to be usable in Microseoft Excel and also to simplify the work of creating visual representations of
combined heart rate data and observation data in Microsoft Excel.
To convert  from a .fit file (Garmin format) to .xml the script uses a free software application called GPSBabel. The
application is available for download here: https://www.gpsbabel.org/. After download, store the application locally.

FUNTIONALITY
Whether run directly from the script or from a compiled executable file (.exe), the script will do the following steps
	INITIALIZING
	- 	Check if the user has previously given path names for folders to 1) look for .fit files, 2) store
	intermediate .xml files and 3) store the final Excel-files (.xlsx). If these folders do not exist, the
	user will be prompted to name their paths to them or create them. The paths will be stored in aconfig file (txt)
 	-  	Check whether the path to GPSBabel.exe is given
	TRANSFORMING
 	- 	Prompt the user to choose the .fit file to be transformed and to give the resulting Excel-file a name
	- 	Run GPSBabel on the selected .fit file and store the resulting .xml file in the designated folder
	DATA TRIMMING
	-	Open the .xml file in Excel and store the file as .xlsx
	-	Remove all data that are not time or heart rate
	CREATE TABLES AND FORMULAE
	- 	Create a sheet for writing (general) frequently observed behaviors and activities
	-	Create a sheet for writing in registered observations
	-	Create a sheet for viewing heart rate and observations for the full registration period in a chart
	-	Create a sheet for viewing heart rate and observations for a selected range of the registration period in a chart
	- 	Modify the sheet that contains heart rate and time with formulae necessary to create charts

CONDITIONS FOR USE AND KNOWN ISSUES
This script has been used and has worked to our satisfaction under the following conditions:
 - .fit files generated using Garmin Forerunner 235 wrist watch
 - PC with Windows 10
 - Microsoft Excel v2003, English language edition.

The script does not support .fit files that spans over more than one date. E.g. if the recording starts
23:00 and ends 00:30 it will crash and the Excel file will not be usable.
In the event of script crash, the original .fit file will not be damaged.
