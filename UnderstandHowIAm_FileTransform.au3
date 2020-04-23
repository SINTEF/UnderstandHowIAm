;~ ==#INFORMATION ABOUT THIS SCRIPT# =====================================================================================
;~ BACKGROUND
;~ This script has been developed as part of the prosject "Forstå hvordan jeg har det" (Understand How I am) 2017-2020.
;~ In the project, heart rate monitors have been used by persons with profound intellectual disability and severe speech,
;~ language or communications difficulties to better understand if and how heart rate can provide more insight for caregivers
;~ into whether the person with severe communication disabilities feels pain, anxiety, stress or other strong emotions as joy.
;~ As part of the project, caregivers have observed behaviors (e.g. a scream or kick) and activities (e.g. meals,
;~ transportation, or physiotherapy sessions) and when they occurred.
;~
;~ PURPOSE
;~ The purpose of this script is to simplify the process of converting files from Garmin heart rate monitor (stored in the
;~ .fit format) to be usable in Microseoft Excel and also to simplify the work of creating visual representations of
;~ combined heart rate data and observation data in Microsoft Excel.
;~ To convert  from a .fit file (Garmin format) to .xml the script uses a free software application called GPSBabel. The
;~ application is available for download here: https://www.gpsbabel.org/. After download, store the application locally.
;~
;~ FUNTIONALITY
;~ Whether run directly from the script or from a compiled executable file (.exe), the script will do the following steps
;~ 		INITIALIZING
;~		- 	Check if the user has previously given path names for folders to 1) look for .fit files, 2) store
;~		intermediate .xml files and 3) store the final Excel-files (.xlsx). If these folders do not exist, the
;~		user will be prompted to name their paths to them or create them. The paths will be stored in aconfig file (txt)
;~  	-  	Check whether the path to GPSBabel.exe is given
;~		TRANSFORMING
;~  	- 	Prompt the user to choose the .fit file to be transformed and to give the resulting Excel-file a name
;~ 		- 	Run GPSBabel on the selected .fit file and store the resulting .xml file in the designated folder
;~		DATA TRIMMING
;~		-	Open the .xml file in Excel and store the file as .xlsx
;~		-	Remove all data that are not time or heart rate
;~		CREATE TABLES AND FORMULAE
;~		- 	Create a sheet for writing (general) frequently observed behaviors and activities
;~		-	Create a sheet for writing in registered observations
;~		-	Create a sheet for viewing heart rate and observations for the full registration period in a chart
;~		-	Create a sheet for viewing heart rate and observations for a selected range of the registration period in a chart
;~ 		- 	Modify the sheet that contains heart rate and time with formulae necessary to create charts
;~
;~ CONDITIONS FOR USE AND KNOWN ISSUES
;~ This script has been used and has worked to our satisfaction under the following conditions:
;~  - .fit files generated using Garmin Forerunner 235 wrist watch
;~  - PC with Windows 10
;~  - Microsoft Excel v2003, English language edition.
;~
;~ The script does not support .fit files that spans over more than one date. E.g. if the recording starts
;~ 23:00 and ends 00:30 it will crash and the Excel file will not be usable.
;~ In the event of script crash, the original .fit file will not be damaged.
;~ =========================================================================================================================


#include <Constants.au3>
#include <GUIConstantsEx.au3>
#include <ColorConstantS.au3>
#include <StaticConstants.au3>
#include <FileConstants.au3>
#include <ButtonConstants.au3>
#include <WindowsConstants.au3>
#include <Array.au3>
#include <File.au3>
#include <Excel.au3>
#include <ExcelConstants.au3>
#include <ExcelChart.au3>
#include <ExcelChartConstants.au3>


Opt("GUIOnEventMode", 1)

Global $g_idExit, $g_oInputField, $g_oOutputField, $g_oOutputFileNameLabel
Global $g_oRawFolderField, $g_oXmlFolderField, $g_oExcelFolderField, $g_oGPSBabelFolderField
Global $g_sChosenLanguage = "English" ;"Norsk" If anything else than "Norsk" is chosen, English language is used

Global $g_aMainPOS, $g_aConfigPOS, $g_oMain_GUI, $g_oConfig_GUI, $g_oChooseColumn_GUI

Global $g_sCurrentFileName = ""
Global $g_sIniFilename = ""
Global $g_sIniSection = "Folders"
Global $g_sIniKey_Raw = "FIT-files folder"
Global $g_sIniKey_XML = "XML-files folder"
Global $g_sIniKey_Excel = "Excel-files folder"
Global $g_sIniKey_GPSBabel = "GPSBabel folder"

Global $g_sIniValue_RawFolder = ""
Global $g_sIniValue_XMLFolder = ""
Global $g_sIniValue_ExcelFolder = ""
Global $g_sIniValue_GPSBabelFolder = ""

Global $g_sLabel_Header_Text = ""
Global $g_sSetConfigurationBtn_Text = ""
Global $g_sChooseInputFileBtn_Text = ""
Global $g_sLabel_Define_Text =  ""
Global $g_sChosenFileLabel_Text = ""
Global $g_sCreateExcelBtn_Text = ""

Global $g_iPrivateColor_DARK_BLUE = 0xC47244 ; 0x4472C4
Global $g_iPrivateColor_MEDIUM_BLUE = 0xDBAA8E ;0x8EAADB
Global $g_iPrivateColor_LIGHT_BLUE = 0xF2E1D9 ;0xD9E1F2

Global $g_iPrivateColor_DARK_YELLOW = 0x00C0FF ;0xFFC000
Global $g_iPrivateColor_MEDIUM_YELLOW = 0x79DFFF ;0xFFDF79
Global $g_iPrivateColor_LIGHT_YELLOW = 0xCCF2FF ;0xFFF2CC

Global $g_iPrivateColor_DARK_GREEN = 0x47AD70 ;0x70AD47
Global $g_iPrivateColor_MEDIUM_GREEN = 0x91CEA9 ;0xA9CE91
Global $g_iPrivateColor_LIGHT_GREEN = 0xDAEFE2 ;0xE2EFDA

Global $g_iNumberOfActivities = 60 ;30
Global $g_iNumberOfEvents = 80 ;50
Global $g_iNumberOfSecondsInSecondInterval = 1

Global $g_iProgressCounter = 0
Global $g_iProgressMaxCount = 47+2*($g_iNumberOfActivities+$g_iNumberOfEvents)

Global $g_iHourMin = 0
Global $g_iMinuteMin = 0
Global $g_iSecondMin = 0

Global $g_iHourMax = 0
Global $g_iMinuteMax = 0
Global $g_iSecondMax = 0

Global $g_iNumberOfMinutes = 0
Global $g_iNumberOfSeconds = 0

Global $g_iS4SelectedTo_Hour = 0
Global $g_iS4SelectedTo_Minute = 0
Global $g_iS4SelectedTo_Second = 0
Global $g_iS4MaxLengthOfShortGraph = 600
Global $g_iLength_OfShortGraph = 0

;Worksheets
Global $g_oExcel, $g_oWorkbook
Global $g_sNameOfSheet_5_CalculationsMainGraph = ""
Global $g_sNameOfSheet_4_GraphSelection = ""
Global $g_sNameOfSheet_3_GraphMain = ""
Global $g_sNameOfSheet_2_Observations = ""
Global $g_sNameOfSheet_1_Descriptions = ""

Global $g_iNumberOfSheets = 5

;Column and Row names on sheet:  $g_sNameOfSheet_1_Descriptions
Global $g_iS1HeaderRow = 1
Global $g_iS1HeaderColumn = 1
Global $g_iS1ReduceVariationWarningRow = $g_iS1HeaderRow +2
Global $g_iS1ReduceVariationExampleRow = $g_iS1ReduceVariationWarningRow +1
Global $g_iS1ReduceVariationWarningColumn = $g_iS1HeaderColumn
Global $g_iS1ReduceVariationExampleColumn = $g_iS1HeaderColumn

Global $g_bS1ShowReduceVariationWarning = True
Global $g_sS1ReduceVariationWarning = ""
Global $g_sS1ReduceVariationExample = ""


Global $g_iS1ActivityFirstHeaderRow = $g_iS1ReduceVariationExampleRow + 2
Global $g_iS1ActivitySecondHeaderRow = $g_iS1ActivityFirstHeaderRow + 2
Global $g_iS1ActivityFirstContentRow = $g_iS1ActivitySecondHeaderRow +1
Global $g_iS1ActivityFirstColumn = 1
Global $g_iS1ActivityDescriptionColumn = $g_iS1ActivityFirstColumn +1
Global $g_iS1ActivityTable_NumberOfColumns = 2
Global $g_iS1NumberOfActivities = 20
Global $g_iS1LastActivityContentRow = $g_iS1ActivityFirstContentRow + $g_iS1NumberOfActivities - 1

Global $g_iS1EventFirstHeaderRow = $g_iS1ReduceVariationExampleRow + 2
Global $g_iS1EventSecondHeaderRow = $g_iS1EventFirstHeaderRow + 2
Global $g_iS1EventFirstContentRow = $g_iS1EventSecondHeaderRow +1
Global $g_iS1EventFirstColumn = $g_iS1ActivityFirstColumn + $g_iS1ActivityTable_NumberOfColumns + 2
Global $g_iS1EventDescriptionColumn = $g_iS1EventFirstColumn +1
Global $g_iS1EventTable_NumberOfColumns = 2
Global $g_iS1NumberOfEvents = 15
Global $g_iS1LastEventContentRow = $g_iS1EventFirstContentRow + $g_iS1NumberOfEvents - 1



;Column and Row names on sheet:  $g_sNameOfSheet_2_Observations
Global $g_iS2TimeSummaryHeaderRow = 1
Global $g_iS2TimeSummaryHeaderColumn = 1
Global $g_iS2TimeSummaryContentRow = $g_iS2TimeSummaryHeaderRow + 1

Global $g_iS2TimeSummaryFromColumn = 3
Global $g_iS2TimeSummaryToColumn = 4
Global $g_iS2TimeSummmaryTableNumberOfColumns = $g_iS2TimeSummaryToColumn

Global $g_iS2HourCorrectionRow = $g_iS2TimeSummaryContentRow + 1
Global $g_iS2HourCorrectionInstructionColumn = 1
Global $g_iS2HourCorrectionInputColumn = $g_iS2TimeSummaryToColumn

Global $g_bS2ShowReduceVariationWarning = $g_bS1ShowReduceVariationWarning
Global $g_sS2ReduceVariationWarning = $g_sS1ReduceVariationWarning
Global $g_iS2ReduceVariationWarningColumn = $g_iS2TimeSummaryToColumn + 3
Global $g_iS2ReduceVariationWarningRow = $g_iS2TimeSummaryHeaderRow



Global $g_iS2ActivityFirstColumn = 1
Global $g_iS2ActivityDescriptionColumn = 2
Global $g_iS2ActivityFromColumn = 3
Global $g_iS2ActivityToColumn = 4
Global $g_iS2ActivityFirstHeaderRow = $g_iS2HourCorrectionRow + 2
Global $g_iS2ActivitySecondHeaderRow = $g_iS2ActivityFirstHeaderRow + 2
Global $g_iS2ActivityFirstContentRow = $g_iS2ActivitySecondHeaderRow + 1
Global $g_iS2LastActivityContentRow = $g_iS2ActivityFirstContentRow + $g_iNumberOfActivities - 1
Global $g_iS2ActivityTable_NumberOfColumns = 4

Global $g_iS2EventFirstColumn = $g_iS2ActivityFirstColumn + $g_iS2ActivityTable_NumberOfColumns + 2
Global $g_iS2EventDescriptionColumn = $g_iS2EventFirstColumn + 1
Global $g_iS2EventFromColumn = $g_iS2EventDescriptionColumn + 1
Global $g_iS2EventToColumn = $g_iS2EventFromColumn + 1

Global $g_iS2EventFirstHeaderRow = $g_iS2ActivityFirstHeaderRow
Global $g_iS2EventSecondHeaderRow = $g_iS2EventFirstHeaderRow + 2
Global $g_iS2EventFirstContentRow = $g_iS2EventSecondHeaderRow + 1
Global $g_iS2LastEventContentRow = $g_iS2EventFirstContentRow + $g_iNumberOfEvents - 1
Global $g_iS2EventTable_NumberOfColumns = 4

Global $g_iS2HourCorrectionValue = 1

;Column and Row names on sheet:  $g_sNameOfSheet_3_GraphMain
Global $g_bS3ShowResolutionWarning = False
Global $g_bS3ShowOverallAverageInChart = True

Global $g_iS3TimeSummaryHeaderRow = 1
Global $g_iS3TimeSummaryHeaderColumn = 1
Global $g_iS3TimeSummaryContentRow = $g_iS3TimeSummaryHeaderRow + 1

Global $g_iS3TimeSummaryFromColumn = 4
Global $g_iS3TimeSummaryToColumn = 5

Global $g_iS3GraphTimeResolutionRow = $g_iS3TimeSummaryContentRow + 1
Global $g_iS3GraphTimeResolutionIntroColumn = $g_iS3TimeSummaryHeaderColumn
Global $g_iS3GraphTimeResolutionInputColumn = $g_iS3TimeSummaryFromColumn
Global $g_iS3GraphTimeResolutionWarningColumn = $g_iS3TimeSummaryToColumn + 1
Global $g_iS3TimeSummmaryTableNumberOfColumns = $g_iS3GraphTimeResolutionWarningColumn

Global $g_iS3GraphNumberOfDatapointsRow = $g_iS3GraphTimeResolutionRow + 1
Global $g_iS3GraphNumberOfDatapointsIntroColumn = $g_iS3TimeSummaryHeaderColumn
Global $g_iS3GraphNumberOfDatapointsValueColumn = $g_iS3TimeSummaryFromColumn
Global $g_iS3GraphNumberOfDatapointsText2Column = $g_iS3TimeSummaryToColumn

Global $g_iS3PulseSummaryHeaderRow = $g_iS3GraphNumberOfDatapointsRow +2
Global $g_iS3PulseSummaryAverageRow = $g_iS3PulseSummaryHeaderRow +1
Global $g_iS3PulseSummaryStDevRow = $g_iS3PulseSummaryAverageRow +1
Global $g_iS3PulseSummaryAvPlusOneStDevRow = $g_iS3PulseSummaryStDevRow +1
Global $g_iS3PulseSummaryAvPlusTwoStDevRow = $g_iS3PulseSummaryAvPlusOneStDevRow +1
Global $g_iS3PulseSummaryIntroColumn = $g_iS3TimeSummaryHeaderColumn
Global $g_iS3PulseSummaryValueColumn = $g_iS3TimeSummaryToColumn

If $g_bS3ShowOverallAverageInChart = True Then
	Global $g_iS3EventHeaderRow = $g_iS3PulseSummaryAvPlusTwoStDevRow + 2
Else
	Global $g_iS3EventHeaderRow = $g_iS3GraphNumberOfDatapointsRow +2
EndIf

Global $g_iS3EventLevelRow = $g_iS3EventHeaderRow + 1
Global $g_iS3EventDistanceRow = $g_iS3EventLevelRow + 1
Global $g_iS3EventIntroColumn = $g_iS3TimeSummaryHeaderColumn
Global $g_iS3EventValueColumn = $g_iS3TimeSummaryFromColumn
Global $g_iS3Event_FirstLevelValue = 155
Global $g_iS3Event_DeltaLevelValue = 3

Global $g_iS3ActivityHeaderRow = $g_iS3EventDistanceRow + 2
Global $g_iS3ActivityLevelRow = $g_iS3ActivityHeaderRow + 1
Global $g_iS3ActivityDistanceRow = $g_iS3ActivityLevelRow + 1
Global $g_iS3ActivityIntroColumn = $g_iS3TimeSummaryHeaderColumn
Global $g_iS3ActivityValueColumn = $g_iS3TimeSummaryFromColumn
Global $g_iS3Activity_FirstLevelValue = 20
Global $g_iS3Activity_DeltaLevelValue = 5

Global $g_iS3GraphFirstRow = $g_iS3GraphNumberOfDatapointsRow +2
Global $g_iS3GraphFirstColumn = $g_iS3PulseSummaryValueColumn +2

;Column and Row names on sheet:  $g_sNameOfSheet_4_GraphSelection
Global $g_bS4ShowResolutionWarning = False
Global $g_bS4ShowOverallAverageInChart = True

Global $g_iS4TimeSummaryHeaderRow = 1
Global $g_iS4TimeSummaryHeaderColumn = 1
Global $g_iS4TimeSummaryContentRow = $g_iS4TimeSummaryHeaderRow + 1

Global $g_iS4TimeSummaryFromColumn = 4
Global $g_iS4TimeSummaryToColumn = 5

Global $g_iS4GraphTimeSelectionRow = $g_iS4TimeSummaryContentRow + 1
Global $g_iS4GraphTimeSelectionIntroColumn = $g_iS4TimeSummaryHeaderColumn
Global $g_iS4GraphTimeSelectionFromColumn = $g_iS4TimeSummaryFromColumn
Global $g_iS4GraphTimeSelectionToColumn = $g_iS4TimeSummaryToColumn

Global $g_iS4GraphTimeResolutionRow = $g_iS4GraphTimeSelectionRow + 1
Global $g_iS4GraphTimeResolutionIntroColumn = $g_iS4TimeSummaryHeaderColumn
Global $g_iS4GraphTimeResolutionInputColumn = $g_iS4TimeSummaryFromColumn
Global $g_iS4GraphTimeResolutionWarningColumn = $g_iS4TimeSummaryToColumn + 1
Global $g_iS4TimeSummmaryTableNumberOfColumns = $g_iS4GraphTimeResolutionWarningColumn

Global $g_iS4GraphNumberOfDatapointsRow = $g_iS4GraphTimeResolutionRow + 1
Global $g_iS4GraphNumberOfDatapointsIntroColumn = $g_iS4TimeSummaryHeaderColumn
Global $g_iS4GraphNumberOfDatapointsValueColumn = $g_iS4TimeSummaryFromColumn
Global $g_iS4GraphNumberOfDatapointsText2Column = $g_iS4TimeSummaryToColumn

Global $g_iS4PulseSummaryHeaderRow = $g_iS4GraphNumberOfDatapointsRow +2
Global $g_iS4PulseSummaryHeaderExtraInfoRow = $g_iS4PulseSummaryHeaderRow +1
Global $g_iS4PulseSummaryAverageRow = $g_iS4PulseSummaryHeaderExtraInfoRow +1
Global $g_iS4PulseSummaryStDevRow = $g_iS4PulseSummaryAverageRow +1
Global $g_iS4PulseSummaryAvPlusOneStDevRow = $g_iS4PulseSummaryStDevRow +1
Global $g_iS4PulseSummaryAvPlusTwoStDevRow = $g_iS4PulseSummaryAvPlusOneStDevRow +1
Global $g_iS4PulseSummaryIntroColumn = $g_iS4TimeSummaryHeaderColumn
Global $g_iS4PulseSummaryValueColumn = $g_iS4TimeSummaryToColumn

If $g_bS4ShowOverallAverageInChart = True Then
	Global $g_iS4EventHeaderRow = $g_iS4PulseSummaryAvPlusTwoStDevRow + 2
Else
	Global $g_iS4EventHeaderRow = $g_iS4GraphNumberOfDatapointsRow + 2
EndIf

Global $g_iS4EventLevelRow = $g_iS4EventHeaderRow + 1
Global $g_iS4EventDistanceRow = $g_iS4EventLevelRow + 1
Global $g_iS4EventIntroColumn = $g_iS4TimeSummaryHeaderColumn
Global $g_iS4EventValueColumn = $g_iS4TimeSummaryFromColumn
Global $g_iS4Event_FirstLevelValue = 155
Global $g_iS4Event_DeltaLevelValue = 3

Global $g_iS4ActivityHeaderRow = $g_iS4EventDistanceRow + 2
Global $g_iS4ActivityLevelRow = $g_iS4ActivityHeaderRow + 1
Global $g_iS4ActivityDistanceRow = $g_iS4ActivityLevelRow + 1
Global $g_iS4ActivityIntroColumn = $g_iS4TimeSummaryHeaderColumn
Global $g_iS4ActivityValueColumn = $g_iS4TimeSummaryFromColumn
Global $g_iS4Activity_FirstLevelValue = 20
Global $g_iS4Activity_DeltaLevelValue = 5

Global $g_sS4AxisWarning = ""
Global $g_iS4AxisWarningColumn = $g_iS4TimeSummaryToColumn + 5
Global $g_iS4AxisStartColumn = $g_iS4AxisWarningColumn
Global $g_iS4AxisStopColumn = $g_iS4AxisStartColumn +1
Global $g_sS4AxisWarningRow = $g_iS4TimeSummaryHeaderRow
Global $g_sS4AxisInfoRow = $g_sS4AxisWarningRow +1
Global $g_sS4AxisValueRow = $g_sS4AxisInfoRow +1

Global $g_iS4GraphFirstRow = $g_iS4GraphNumberOfDatapointsRow +2
Global $g_iS4GraphFirstColumn = $g_iS4PulseSummaryValueColumn +2



;Column and Row names on sheet:  $g_sNameOfSheet_5_CalculationsMainGraph
Global $g_sS5MainTable_Name =  "MainPulseTable"
Global $g_iS5FirstHeaderRow = 1
Global $g_iS5Column_Date = 1
Global $g_iS5Column_Hour = $g_iS5Column_Date + 1
Global $g_iS5Column_CorrectedHour = $g_iS5Column_Hour + 1
Global $g_iS5Column_Minute = $g_iS5Column_CorrectedHour + 1
Global $g_iS5Column_Second = $g_iS5Column_Minute + 1

Global $g_iS5Column_CorrectedTime = $g_iS5Column_Second + 1
Global $g_sS5CorrectedTimeForMainTable_ColumnName = ""
Global $g_iS5Column_Pulse = $g_iS5Column_CorrectedTime + 1
Global $g_sS5PulseForMainTable_ColumnName = ""

Global $g_sS5SecondsTable_Name  = "SecondsTable"
Global $g_iS5Column_SecondInterval = $g_iS5Column_Pulse + 1
Global $g_sS5SecondInterval_ColumnName = ""

;Long graph Columns
Global $g_iS5Column_TimeForLongGraph = $g_iS5Column_SecondInterval + 1
Global $g_sS5TimeForLongGraph_ColumnName = ""
Global $g_iS5Column_PulseForLongGraph = $g_iS5Column_TimeForLongGraph + 1
Global $g_sS5PulseForLongGraph_ColumnName = ""

;Short graph Columns
Global $g_iS5Column_TimeForShortGraph = $g_iS5Column_PulseForLongGraph + 1
Global $g_sS5TimeForShortGraph_ColumnName = ""
Global $g_iS5Column_PulseForShortGraph = $g_iS5Column_TimeForShortGraph + 1
Global $g_sS5PulseForShortGraph_ColumnName = ""

;Long Graph summary
Global $g_iS5Column_LongGraphSummary_Header = $g_iS5Column_PulseForShortGraph +2
Global $g_iS5Column_LongGraphSummary_Value = $g_iS5Column_LongGraphSummary_Header +1
Global $g_iS5Row_LongGraphSummary_Header = $g_iS5FirstHeaderRow
Global $g_iS5Row_LongGraphAverage_Header = $g_iS5Row_LongGraphSummary_Header +1
Global $g_iS5Row_LongGraphAverage_FirstValue = $g_iS5Row_LongGraphAverage_Header +1
Global $g_iS5Row_LongGraphAverage_LastValue = $g_iS5Row_LongGraphAverage_FirstValue +1

Global $g_iS5Row_LongGraphOneStDev_Header = $g_iS5Row_LongGraphAverage_LastValue +2
Global $g_iS5Row_LongGraphOneStDev_FirstValue = $g_iS5Row_LongGraphOneStDev_Header +1
Global $g_iS5Row_LongGraphOneStDev_LastValue = $g_iS5Row_LongGraphOneStDev_FirstValue +1

Global $g_iS5Row_LongGraphTwoStDev_Header = $g_iS5Row_LongGraphOneStDev_LastValue +2
Global $g_iS5Row_LongGraphTwoStDev_FirstValue = $g_iS5Row_LongGraphTwoStDev_Header +1
Global $g_iS5Row_LongGraphTwoStDev_LastValue = $g_iS5Row_LongGraphTwoStDev_FirstValue +1

;Short graph summary
Global $g_iS5Column_ShortGraphSummary_Header = $g_iS5Column_LongGraphSummary_Value +2
Global $g_iS5Column_ShortGraphSummary_Value = $g_iS5Column_ShortGraphSummary_Header + 2

Global $g_iS5Row_ShortGraphSummary_Header = $g_iS5FirstHeaderRow
Global $g_iS5Row_ShortGraphAverage_Header = $g_iS5Row_ShortGraphSummary_Header +1
Global $g_iS5Row_ShortGraphAverage_FirstValue = $g_iS5Row_ShortGraphAverage_Header +1
Global $g_iS5Row_ShortGraphAverage_LastValue = $g_iS5Row_ShortGraphAverage_FirstValue +1

Global $g_iS5Row_ShortGraphOneStDev_Header = $g_iS5Row_ShortGraphAverage_LastValue +2
Global $g_iS5Row_ShortGraphOneStDev_FirstValue = $g_iS5Row_ShortGraphOneStDev_Header +1
Global $g_iS5Row_ShortGraphOneStDev_LastValue = $g_iS5Row_ShortGraphOneStDev_FirstValue +1

Global $g_iS5Row_ShortGraphTwoStDev_Header = $g_iS5Row_ShortGraphOneStDev_LastValue +2
Global $g_iS5Row_ShortGraphTwoStDev_FirstValue = $g_iS5Row_ShortGraphTwoStDev_Header +1
Global $g_iS5Row_ShortGraphTwoStDev_LastValue = $g_iS5Row_ShortGraphTwoStDev_FirstValue +1
Global $g_iS5Row_ShortGraphFirstValidRow = $g_iS5Row_ShortGraphTwoStDev_LastValue +2
Global $g_iS5Row_ShortGraphFirstValidCell = $g_iS5Row_ShortGraphFirstValidRow + 1

Global $g_iS5MainHeaderRow = 1
Global $g_iS5FirstContentRow = $g_iS5MainHeaderRow + 1
Global $g_iS5LastContentRow = $g_iS5FirstContentRow
Global $g_iS5LastMinuteIntervalRow = $g_iS5FirstContentRow
Global $g_iS5LastSecondIntervalRow = $g_iS5FirstContentRow

;For long graph
Global $g_iS5Activity_ForLong_HeaderColumn = $g_iS5Column_ShortGraphSummary_Value +3
Global $g_iS5Activity_ForLong_TimeColumn = $g_iS5Activity_ForLong_HeaderColumn + 1
Global $g_iS5Activity_ForLong_ValueColumn = $g_iS5Activity_ForLong_TimeColumn+1
Global $g_iS5Activity_ForLong_FirstHeaderRow = $g_iS5MainHeaderRow
Global $g_iS5Activity_ForLong_FirstContentRow = $g_iS5Activity_ForLong_FirstHeaderRow +1
Global $g_iS5Activity_ForLong_NumberOfRowsInSubTable  = 4 ;(Includes 1 space row

Global $g_iS5Event_ForLong_HeaderColumn = $g_iS5Activity_ForLong_ValueColumn +3
Global $g_iS5Event_ForLong_TimeColumn = $g_iS5Event_ForLong_HeaderColumn +1
Global $g_iS5Event_ForLong_ValueColumn = $g_iS5Event_ForLong_TimeColumn+1
Global $g_iS5Event_ForLong_FirstHeaderRow = $g_iS5MainHeaderRow
Global $g_iS5Event_ForLong_FirstContentRow = $g_iS5Event_ForLong_FirstHeaderRow +1
Global $g_iS5Event_ForLong_NumberOfRowsInSubTable  = 4 ;(Includes 1 space row

;For short graph
Global $g_iS5Activity_ForShort_HeaderColumn = $g_iS5Event_ForLong_ValueColumn +3 +1
Global $g_iS5Activity_ForShort_TimeColumn = $g_iS5Activity_ForShort_HeaderColumn + 1
Global $g_iS5Activity_ForShort_ValueColumn = $g_iS5Activity_ForShort_TimeColumn+1
Global $g_iS5Activity_ForShort_FirstHeaderRow = $g_iS5MainHeaderRow
Global $g_iS5Activity_ForShort_FirstContentRow = $g_iS5Activity_ForShort_FirstHeaderRow +1
Global $g_iS5Activity_ForShort_NumberOfRowsInSubTable  = 4 ;(Includes 1 space row

Global $g_iS5Event_ForShort_HeaderColumn = $g_iS5Activity_ForShort_ValueColumn +3
Global $g_iS5Event_ForShort_TimeColumn = $g_iS5Event_ForShort_HeaderColumn +1
Global $g_iS5Event_ForShort_ValueColumn = $g_iS5Event_ForShort_TimeColumn+1
Global $g_iS5Event_ForShort_FirstHeaderRow = $g_iS5MainHeaderRow
Global $g_iS5Event_ForShort_FirstContentRow = $g_iS5Event_ForShort_FirstHeaderRow +1
Global $g_iS5Event_ForShort_NumberOfRowsInSubTable  = 4 ;(Includes 1 space row

Global $g_iTableMainHeaderColorActivity = $g_iPrivateColor_DARK_GREEN
Global $g_iTableSecondHeaderColorActivity = $g_iPrivateColor_MEDIUM_GREEN
Global $g_iTableContentColorActivity = $g_iPrivateColor_LIGHT_GREEN

Global $g_iTableMainHeaderColorEvent = $g_iPrivateColor_DARK_YELLOW
Global $g_iTableSecondHeaderColorEvent = $g_iPrivateColor_MEDIUM_YELLOW
Global $g_iTableContentColorEvent = $g_iPrivateColor_LIGHT_YELLOW

Global $g_iTableMainHeaderColorMaster = $g_iPrivateColor_DARK_BLUE
Global $g_iTableSecondHeaderColorMaster = $g_iPrivateColor_MEDIUM_BLUE
Global $g_iTableContentColorMaster = $g_iPrivateColor_LIGHT_BLUE

Global $g_iTableMainHeaderFontSize = 18
Global $g_iTableSecondHeaderFontSize = 14
Global $g_iTableContentHeaderFontSize = 11

Global $g_iTableMainHeaderFontColor = $COLOR_WHITE
Global $g_iTableSecondHeaderFontColor = $COLOR_WHITE
Global $g_iTableContentFontColor = $COLOR_BLACK
Global $g_iTableOutsideBorderThickness = $xlThick
Global $g_iTableInsideBorderThickness = $xlThin

;xlBordersIndex
Global Const $g_xlDiagonalDown = 5
Global Const $g_xlDiagonalUp = 6
Global Const $g_xlEdgeLeft = 7
Global Const $g_xlEdgeTop = 8
Global Const $g_xlEdgeBottom = 9
Global Const $g_xlEdgeRight = 10
Global Const $g_xlInsideVertical = 11
Global Const $g_xlInsideHorizontal = 12

;xlBorderWeightEnumeration
;Global Const $xlHairline = 1
;Global Const $xlThin = 2
;Global Const $xlThick = 4
;Global Const $xlMedium = -4138

;xlLineStyle Enumeration
Global Const $g_xlContinuous = 1 ;Continuous line.
Global Const $g_xlDash = -4115 ;Dashed line.
Global Const $g_xlDashDot = 4 ;Alternating dashes and dots.
Global Const $g_xlDashDotDot = 5 ;Dash followed by two dots.
Global Const $g_xlDot = -4118 ;Dotted line.
Global Const $g_xlDouble = -4119 ;Double line.
Global Const $g_xlLineStyleNone = -4142 ;No line.
Global Const $g_xlSlantDashDot = 13 ;Slanted dashes.

;Constants for GUIs
Global $g_iFirstColumn_left = 20
Global $g_iButtonHeight = 20
Global $g_iButtonWidth = 90
Global $g_iConfigButtonWidth = 110
Global $g_iCtrlDistanceH = 10
Global $g_iCtrlDistanceV = 20

Global $g_iFirstRowTop = 10
Global $g_iSecondRowTop = $g_iFirstRowTop + $g_iButtonHeight + $g_iCtrlDistanceV
Global $g_iThirdRowTop = $g_iSecondRowTop + $g_iButtonHeight + $g_iCtrlDistanceV
Global $g_iFourthRowTop = $g_iThirdRowTop + $g_iButtonHeight + $g_iCtrlDistanceV
Global $g_iFifthRowTop = $g_iFourthRowTop + $g_iButtonHeight + $g_iCtrlDistanceV * 2


_Main()

Func _Main()

	Local $l_oCreateExcelBtn, $l_oSetConfigurationBtn, $l_oChooseInputFileBtn, $l_oChooseOutputFileBtn, $l_oOutputFileExtensionLabel, $l_oChosenFileLabel
	Local $l_CtrlPlacementH

	Load_Language()

	$g_oMain_GUI = GUICreate($g_sLabel_Header_Text, 550, $g_iFifthRowTop + $g_iButtonHeight + $g_iCtrlDistanceV)

	GUICtrlCreateLabel($g_sLabel_Header_Text, $g_iFirstColumn_left, $g_iFirstRowTop)
	;$l_oSetConfigurationBtn = GUICtrlCreateButton("", 500, $g_iFirstRowTop, $g_iButtonHeight + 20, $g_iButtonHeight + 20, $BS_BITMAP)
	;GUICtrlSetImage($l_oSetConfigurationBtn, @ScriptDir & "\settings_icon_small_bmp.bmp")
	$l_oSetConfigurationBtn = GUICtrlCreateButton($g_sSetConfigurationBtn_Text,550-$g_iButtonWidth-20, $g_iFirstRowTop,$g_iButtonWidth, $g_iButtonHeight)
	GUICtrlSetOnEvent($l_oSetConfigurationBtn, "OnSetConfiguration")

	$l_CtrlPlacementH = $g_iFirstColumn_left
	$l_oChooseInputFileBtn = GUICtrlCreateButton($g_sChooseInputFileBtn_Text , $l_CtrlPlacementH, $g_iSecondRowTop, $g_iButtonWidth, $g_iButtonHeight)
	$l_CtrlPlacementH = $l_CtrlPlacementH + $g_iButtonWidth + $g_iCtrlDistanceH
	$g_oInputField = GUICtrlCreateInput("", $l_CtrlPlacementH, $g_iSecondRowTop, 350, $g_iButtonHeight)
	GUICtrlSetOnEvent($l_oChooseInputFileBtn, "OnChooseInputFile")

	;$l_oChooseOutputFileBtn = GUICtrlCreateButton("Definer outputfil", $g_iFirstColumn_left, $g_iThirdRowTop, $g_iButtonWidth, $g_iButtonHeight)
	GUICtrlCreateLabel($g_sLabel_Define_Text, $g_iFirstColumn_left, $g_iThirdRowTop, $g_iButtonWidth * 3, $g_iButtonHeight)
	;GUICtrlSetOnEvent($l_oChooseOutputFileBtn, "OnChooseOutputFile")

	Local $l_iOutputColor = $COLOR_White
	$l_CtrlPlacementH = $g_iFirstColumn_left
	$l_oChosenFileLabel = GUICtrlCreateLabel($g_sChosenFileLabel_Text, $l_CtrlPlacementH, $g_iFourthRowTop, $g_iButtonWidth, $g_iButtonHeight, $SS_LEFT)
	$l_CtrlPlacementH = $l_CtrlPlacementH + $g_iButtonWidth + $g_iCtrlDistanceH
	$g_oOutputFileNameLabel = GUICtrlCreateLabel("", $l_CtrlPlacementH, $g_iFourthRowTop + 1, $g_iButtonWidth, $g_iButtonHeight - 2, $SS_RIGHT)
	GUICtrlSetBkColor($g_oOutputFileNameLabel, $l_iOutputColor)
	$l_CtrlPlacementH = $l_CtrlPlacementH + $g_iButtonWidth
	$g_oOutputField = GUICtrlCreateInput("", $l_CtrlPlacementH, $g_iFourthRowTop, 250, $g_iButtonHeight, $SS_LEFT)
	$l_CtrlPlacementH = $l_CtrlPlacementH + 250
	GUICtrlSetBkColor($g_oOutputField, $l_iOutputColor)
	$l_oOutputFileExtensionLabel = GUICtrlCreateLabel(".xlsx", $l_CtrlPlacementH, $g_iFourthRowTop + 1, $g_iButtonWidth * 0.5, $g_iButtonHeight - 2, $SS_LEFT)
	GUICtrlSetBkColor($l_oOutputFileExtensionLabel, $l_iOutputColor)

	$l_CtrlPlacementH = $g_iFirstColumn_left + $g_iButtonWidth
	$l_oCreateExcelBtn = GUICtrlCreateButton($g_sCreateExcelBtn_Text, $l_CtrlPlacementH, $g_iFifthRowTop, $g_iButtonWidth, $g_iButtonHeight)
	GUICtrlSetOnEvent($l_oCreateExcelBtn, "OnCreateExcel")

	$l_CtrlPlacementH = $l_CtrlPlacementH + $g_iButtonWidth + $g_iCtrlDistanceH
	$g_idExit = GUICtrlCreateButton("Exit", $l_CtrlPlacementH, $g_iFifthRowTop, $g_iButtonWidth, $g_iButtonHeight)
	GUICtrlSetOnEvent($g_idExit, "OnExit")

	;$l_CtrlPlacementH = $l_CtrlPlacementH + $g_iButtonWidth + $g_iCtrlDistanceH + $g_iButtonWidth
	;$g_idExit = GUICtrlCreateButton("Lag Excel fra XML", $l_CtrlPlacementH, $g_iFifthRowTop, $g_iButtonWidth, $g_iButtonHeight)
	;GUICtrlSetOnEvent($g_idExit, "OnCreateExcelFromXML")



	GUISetOnEvent($GUI_EVENT_CLOSE, "OnExit", $g_oMain_GUI)

	$g_aMainPOS = WinGetPos($g_oMain_GUI)
	GUISetState(@SW_SHOW, $g_oMain_GUI) ; display the GUI

	If FileExists(@ScriptDir & "\" & $g_sIniFilename) = 0 Then
		If $g_sChosenLanguage = "Norsk" Then
			MsgBox($MB_SYSTEMMODAL, "Konfigureringsfil finnes ikke", "Det finnes ennå ingen konfigureringsfil. Før vi går i gang må vi opprette den.")
		Else
			MsgBox($MB_SYSTEMMODAL, "Config file does not exist", "Could not find any config file. Before we start we must create one.")
		EndIf
		IniWrite(@ScriptDir & "\" & $g_sIniFilename, "Title", "Konfigurering av Forstaa-meg filomformer / Configuration of Understand how I am file transformer", "")
		IniWrite(@ScriptDir & "\" & $g_sIniFilename, $g_sIniSection, $g_sIniKey_Raw, "")
		IniWrite(@ScriptDir & "\" & $g_sIniFilename, $g_sIniSection, $g_sIniKey_XML, "")
		IniWrite(@ScriptDir & "\" & $g_sIniFilename, $g_sIniSection, $g_sIniKey_Excel, "")
		IniWrite(@ScriptDir & "\" & $g_sIniFilename, $g_sIniSection, $g_sIniKey_GPSBabel, "")
		_ReadConfigurationFile()
		_SetConfiguration(1)
	Else
		_ReadConfigurationFile()
	EndIf

	While 1
		Sleep(1000)
	WEnd
EndFunc   ;==>_Main

Func Load_Language()

	If $g_sChosenLanguage = "Norsk" Then
		$g_sIniFilename = "ForstaaMeg_config.ini"
		$g_sLabel_Header_Text = "Forstå meg! - filomformer"
		$g_sSetConfigurationBtn_Text = "Innstillinger"
		$g_sChooseInputFileBtn_Text = "Velg inputfil"
		$g_sLabel_Define_Text =  "Definer et beskrivende navn til excelfilen"
		$g_sChosenFileLabel_Text = "Valgt filnavn:      "
		$g_sCreateExcelBtn_Text = "Lag Excel-fil"

		$g_sNameOfSheet_5_CalculationsMainGraph = "For beregning hovedgraf"
		$g_sNameOfSheet_4_GraphSelection = "Utdrag av graf"
		$g_sNameOfSheet_3_GraphMain = "Oversiktsgraf"
		$g_sNameOfSheet_2_Observations = "Observasjoner"
		$g_sNameOfSheet_1_Descriptions = "Beskrivelse av aktiviteter"

		$g_sS1ReduceVariationWarning = "OBS: Prøv å begrense antallet ulike aktiviteter og hendelser / atferder."
		$g_sS1ReduceVariationExample = "Eksempel: Kan ""rop"" og ""hyl"" slås sammen til én atferd?"

		$g_sS2ReduceVariationWarning = "OBS: Prøv å begrense antallet ulike aktiviteter og hendelser / atferder."

		$g_sS4AxisWarning = "OBS: Når du endrer tiden for utsnittet må du manuelt endre start og stopp for den horisontale aksen i grafen"
		$g_sS5CorrectedTimeForMainTable_ColumnName = "Korrigert tid"
		$g_sS5PulseForMainTable_ColumnName = "Puls"

		$g_sS5SecondInterval_ColumnName = "Sekundintervall"
		$g_sS5TimeForLongGraph_ColumnName = "Tid lang graf"
		$g_sS5PulseForLongGraph_ColumnName = "Puls lang graf"
		$g_sS5TimeForShortGraph_ColumnName = "Tid kort graf"
		$g_sS5PulseForShortGraph_ColumnName = "Puls kort graf"

	Else
		$g_sIniFilename = "UnderstandHowIAm_config.ini"
		$g_sLabel_Header_Text = "Understand how I am! - file transform"
		$g_sSetConfigurationBtn_Text = "Settings"
		$g_sChooseInputFileBtn_Text = "Choose input file"
		$g_sLabel_Define_Text =  "Define a describing name for your excel file"
		$g_sChosenFileLabel_Text = "Chosen file name:      "
		$g_sCreateExcelBtn_Text = "Create Excel file"

		$g_sNameOfSheet_5_CalculationsMainGraph = "Raw data"
		$g_sNameOfSheet_4_GraphSelection = "Graph selection"
		$g_sNameOfSheet_3_GraphMain = "Graph overview"
		$g_sNameOfSheet_2_Observations = "Observations"
		$g_sNameOfSheet_1_Descriptions = "Descriptions of activities"

		$g_sS1ReduceVariationWarning = "OBS: Try to reduce the number of different activities and events / behaviors."
		$g_sS1ReduceVariationExample = "Example: Could ""shout"" and ""scream"" be the decribed as the same as the same behavior?"

		$g_sS2ReduceVariationWarning = "OBS: Try to reduce the number of different activities and events / behaviors."

		$g_sS4AxisWarning = "OBS: When changing the selected time you must manually enter start and stop for the horizontal axis in the chart"
		$g_sS5CorrectedTimeForMainTable_ColumnName = "Corrected time"
		$g_sS5PulseForMainTable_ColumnName = "HR"

		$g_sS5SecondInterval_ColumnName = "Second interval"
		$g_sS5TimeForLongGraph_ColumnName = "Time long graph"
		$g_sS5PulseForLongGraph_ColumnName = "HR long graph"
		$g_sS5TimeForShortGraph_ColumnName = "Time short graph"
		$g_sS5PulseForShortGraph_ColumnName = "HR short graph"
	EndIf

EndFunc

; --------------- Config Functions ---------------
Func AddInfoToFileNameLabel()
	Local $l_sInputFile = GUICtrlRead($g_oInputField)
	Local $l_aFileDateTime = FileGetTime("" & $l_sInputFile & "")
	If @error Then
		GUICtrlSetData($g_oOutputFileNameLabel, "")
		MsgBox($MB_SYSTEMMODAL, "NameLabel Error", "Could not read DateTime from " & "" & $l_sInputFile & "")
	Else
		GUICtrlSetData($g_oOutputFileNameLabel, $l_aFileDateTime[0] & "-" & $l_aFileDateTime[1] & "-" & $l_aFileDateTime[2] & "_" & $l_aFileDateTime[3] & $l_aFileDateTime[4] & "_")
		$g_sCurrentFileName = GUICtrlRead($g_oOutputFileNameLabel) & GUICtrlRead($g_oOutputField)
	EndIf
EndFunc   ;==>AddInfoToFileNameLabel

Func OnSetConfiguration()
	_SetConfiguration(1)
EndFunc   ;==>OnSetConfiguration
Func _SetConfiguration($l_iActiveTab)

	$g_aMainPOS = WinGetPos($g_oMain_GUI)
	Local $l_iFolderNameFieldWidth = 350
	Local $l_iFolderInfoLabelWidth = 320
	Local $l_oTabItemStoreFolders, $l_oTabItemGPSBabelFolder, $l_oConfigurationTab
	Local $l_iGuiWidth = $g_iFirstColumn_left + $g_iConfigButtonWidth + $l_iFolderNameFieldWidth + $g_iCtrlDistanceH * 4 + $l_iFolderInfoLabelWidth
	Local $l_iGuiHeight = $g_iFifthRowTop + $g_iButtonHeight + 10
	Local $l_iTabWidth = $l_iGuiWidth - 20
	Local $l_iTabHeight = $g_iFourthRowTop + $g_iButtonHeight + 10
	Local $l_CtrlPlacementH

	;$g_oConfig_GUI = GUICreate("Konfigurering",550, $g_iFifthRowTop + $g_iButtonHeight + 10, $g_aMainPOS[0] + $g_aMainPOS[2], $g_aMainPOS[1], -1, -1, $g_oMain_GUI)
	$g_oConfig_GUI = GUICreate("Config", $l_iGuiWidth, $l_iGuiHeight, $g_aMainPOS[0], $g_aMainPOS[1], -1, -1, $g_oMain_GUI)
	$l_oConfigurationTab = GUICtrlCreateTab(10, 10, $l_iTabWidth, $l_iTabHeight)

	$l_oTabItemStoreFolders = GUICtrlCreateTabItem("Folders")

	$l_CtrlPlacementH = $g_iFirstColumn_left
	Local $chooseRawFileFolderBtn = GUICtrlCreateButton("Choose folder for .fit", $l_CtrlPlacementH, $g_iSecondRowTop, $g_iConfigButtonWidth, $g_iButtonHeight)
	$l_CtrlPlacementH = $l_CtrlPlacementH + $g_iConfigButtonWidth + $g_iCtrlDistanceH
	$g_oRawFolderField = GUICtrlCreateInput("", $l_CtrlPlacementH, $g_iSecondRowTop, $l_iFolderNameFieldWidth, $g_iButtonHeight)
	GUICtrlSetOnEvent($chooseRawFileFolderBtn, "OnChooseRawFileFolder")
	$l_CtrlPlacementH = $l_CtrlPlacementH + $l_iFolderNameFieldWidth + $g_iCtrlDistanceH
	GUICtrlCreateLabel("Choose a  folder for raw files (.fit) from the Garmin watch", $l_CtrlPlacementH, $g_iSecondRowTop, $l_iFolderInfoLabelWidth)
	GUICtrlSetData($g_oRawFolderField, $g_sIniValue_RawFolder)

	$l_CtrlPlacementH = $g_iFirstColumn_left
	Local $chooseIntermediateFileFolderBtn = GUICtrlCreateButton("Choose folder for .xml", $l_CtrlPlacementH, $g_iThirdRowTop, $g_iConfigButtonWidth, $g_iButtonHeight)
	$l_CtrlPlacementH = $l_CtrlPlacementH + $g_iConfigButtonWidth + $g_iCtrlDistanceH
	$g_oXmlFolderField = GUICtrlCreateInput("", $l_CtrlPlacementH, $g_iThirdRowTop, $l_iFolderNameFieldWidth, $g_iButtonHeight)
	GUICtrlSetOnEvent($chooseIntermediateFileFolderBtn, "OnChooseIntermediateFileFolder")
	$l_CtrlPlacementH = $l_CtrlPlacementH + $l_iFolderNameFieldWidth + $g_iCtrlDistanceH
	GUICtrlCreateLabel("Choose a  folder for converted files (.xml)", $l_CtrlPlacementH, $g_iThirdRowTop, $l_iFolderInfoLabelWidth)
	GUICtrlSetData($g_oXmlFolderField, $g_sIniValue_XMLFolder)

	$l_CtrlPlacementH = $g_iFirstColumn_left
	Local $chooseExcelFileFolderBtn = GUICtrlCreateButton("Choose folder for .xlsx", $g_iFirstColumn_left, $g_iFourthRowTop, $g_iConfigButtonWidth, $g_iButtonHeight)
	$l_CtrlPlacementH = $l_CtrlPlacementH + $g_iConfigButtonWidth + $g_iCtrlDistanceH
	$g_oExcelFolderField = GUICtrlCreateInput("", $l_CtrlPlacementH, $g_iFourthRowTop, $l_iFolderNameFieldWidth, $g_iButtonHeight)
	GUICtrlSetOnEvent($chooseExcelFileFolderBtn, "OnChooseExcelFileFolder")
	$l_CtrlPlacementH = $l_CtrlPlacementH + $l_iFolderNameFieldWidth + $g_iCtrlDistanceH
	GUICtrlCreateLabel("Choose a  folder for Excel files (.xlsx)", $l_CtrlPlacementH, $g_iFourthRowTop, $l_iFolderInfoLabelWidth)
	GUICtrlSetData($g_oExcelFolderField, $g_sIniValue_ExcelFolder)

	$l_oTabItemGPSBabelFolder = GUICtrlCreateTabItem("GPSBabel")

	$l_CtrlPlacementH = $g_iFirstColumn_left
	GUICtrlCreateLabel("Choose the folder where the GPSBabel-application is found (gpsbabel.exe)", $l_CtrlPlacementH, $g_iSecondRowTop, $l_iTabWidth - $l_CtrlPlacementH * 2)
	Local $chooseGPSBabelFolderBtn = GUICtrlCreateButton("Choose folder", $l_CtrlPlacementH, $g_iThirdRowTop, $g_iConfigButtonWidth, $g_iButtonHeight)
	$l_CtrlPlacementH = $l_CtrlPlacementH + $g_iConfigButtonWidth + $g_iCtrlDistanceH
	$g_oGPSBabelFolderField = GUICtrlCreateInput("", $l_CtrlPlacementH, $g_iThirdRowTop, $l_iFolderNameFieldWidth, $g_iButtonHeight)
	GUICtrlSetOnEvent($chooseGPSBabelFolderBtn, "OnChooseGPSBabelFolder")

	GUICtrlSetData($g_oGPSBabelFolderField, $g_sIniValue_GPSBabelFolder)

	GUICtrlCreateTabItem("")
	If ($l_iActiveTab = 1) Then
		GUICtrlSetState($l_oTabItemStoreFolders, $GUI_SHOW)
	Else
		GUICtrlSetState($l_oTabItemGPSBabelFolder, $GUI_SHOW)
	EndIf

	Local $l_oSaveBtn = GUICtrlCreateButton("Save", $g_iFirstColumn_left, $g_iFifthRowTop, $g_iConfigButtonWidth, $g_iButtonHeight)
	GUICtrlSetOnEvent($l_oSaveBtn, "OnSaveConfig")
	Local $l_oCancelBtn = GUICtrlCreateButton("Cancel", $g_iFirstColumn_left + $g_iCtrlDistanceH + $g_iConfigButtonWidth, $g_iFifthRowTop, $g_iConfigButtonWidth, $g_iButtonHeight)
	GUICtrlSetOnEvent($l_oCancelBtn, "OnExitConfiguration")
	Local $l_oFinishBtn = GUICtrlCreateButton("Close config", $g_iFirstColumn_left + ($g_iCtrlDistanceH + $g_iConfigButtonWidth) * 2, $g_iFifthRowTop, $g_iConfigButtonWidth * 2, $g_iButtonHeight)
	GUICtrlSetOnEvent($l_oFinishBtn, "OnFinishConfig")

	GUISetOnEvent($GUI_EVENT_CLOSE, "OnExitConfiguration", $g_oConfig_GUI)

	GUISetState(@SW_DISABLE, $g_oMain_GUI)
	GUISetState(@SW_SHOW, $g_oConfig_GUI)
	WinActivate($g_oConfig_GUI)
EndFunc   ;==>_SetConfiguration


Func OnChooseRawFileFolder()
	; Create a constant variable in Local scope of the message to display in FileSelectFolder.
	Local $l_sMessage1, $l_sMessage2

	If $g_sChosenLanguage = "Norsk" Then
		$l_sMessage1 = "Velg folderen der filene (.fit) fra Garmin-klokka ligger"
		$l_sMessage2 = "Ingen folder ble valgt."
	Else
		$l_sMessage1 = "Choose the folder where you will store the files (.fit) from the Garmin watch"
		$l_sMessage2 = "No folder was chosen."
	EndIf

	; Display an open dialog to select a file.
	Local $sFileSelectFolder = FileSelectFolder($l_sMessage1, @ScriptDir)
	If @error Then
		; Display the error message.
		MsgBox($MB_SYSTEMMODAL, "", $l_sMessage2)
	Else
		; Display the selected folder.
		;MsgBox($MB_SYSTEMMODAL, "", "You chose the following folder:" & @CRLF & $sFileSelectFolder)
		GUICtrlSetData($g_oRawFolderField, $sFileSelectFolder)
	EndIf
EndFunc   ;==>OnChooseRawFileFolder

Func OnChooseIntermediateFileFolder()
	; Create a constant variable in Local scope of the message to display in FileSelectFolder.
	Local $l_sMessage1, $l_sMessage2

	If $g_sChosenLanguage = "Norsk" Then
		$l_sMessage1 = "Velg folderen der transformerte filer (.xml) skal ligge"
		$l_sMessage2 = "Ingen folder ble valgt."
	Else
		$l_sMessage1 = "Choose the folder where the transformed files (.xml) will be stored"
		$l_sMessage2 = "No folder was chosen."
	EndIf

	; Display an open dialog to select a file.
	Local $sFileSelectFolder = FileSelectFolder($l_sMessage1, @ScriptDir)
	If @error Then
		; Display the error message.
		MsgBox($MB_SYSTEMMODAL, "", $l_sMessage2)
	Else
		; Display the selected folder.
		;MsgBox($MB_SYSTEMMODAL, "", "You chose the following folder:" & @CRLF & $sFileSelectFolder)
		GUICtrlSetData($g_oXmlFolderField, $sFileSelectFolder)
	EndIf
EndFunc   ;==>OnChooseIntermediateFileFolder

Func OnChooseExcelFileFolder()
	; Create a constant variable in Local scope of the message to display in FileSelectFolder.
	Local $l_sMessage1, $l_sMessage2

	If $g_sChosenLanguage = "Norsk" Then
		$l_sMessage1 = "Velg folderen de ferdige excel-filene (.xlsx) skal ligge"
		$l_sMessage2 = "Ingen folder ble valgt."
	Else
		$l_sMessage1 = "Choose the folder where the final excel files (.xlsx) will be stored"
		$l_sMessage2 = "No folder was chosen."
	EndIf

	; Display an open dialog to select a file.
	Local $l_sFileSelectFolder = FileSelectFolder($l_sMessage1, @ScriptDir)
	If @error Then
		; Display the error message.
		MsgBox($MB_SYSTEMMODAL, "", $l_sMessage2)
	Else
		; Display the selected folder.
		;MsgBox($MB_SYSTEMMODAL, "", "You chose the following folder:" & @CRLF & $sFileSelectFolder)
		GUICtrlSetData($g_oExcelFolderField, $l_sFileSelectFolder)
	EndIf
EndFunc   ;==>OnChooseExcelFileFolder
Func OnChooseGPSBabelFolder()
	; Create a constant variable in Local scope of the message to display in FileSelectFolder.
	Local $l_sMessage1, $l_sMessage2

	If $g_sChosenLanguage = "Norsk" Then
		$l_sMessage1 = "Velg folderen der applikasjonen gpsbabel.exe ligger. (Folderen heter vanligvis GPSBabel.)"
		$l_sMessage2 = "Ingen folder ble valgt."
	Else
		$l_sMessage1 = "Choose the folder where the application gpsbabel.exe is found. (The folder is normally named GPSBabel.)"
		$l_sMessage2 = "No folder was chosen."
	EndIf
	; Display an open dialog to select a file.
	Local $l_sFileSelectFolder = FileSelectFolder($l_sMessage1, @ProgramFilesDir)
	If @error Then
		; Display the error message.
		MsgBox($MB_SYSTEMMODAL, "", $l_sMessage2)
	Else
		; Display the selected folder.
		;MsgBox($MB_SYSTEMMODAL, "", "You chose the following folder:" & @CRLF & $sFileSelectFolder)
		GUICtrlSetData($g_oGPSBabelFolderField, $l_sFileSelectFolder)
	EndIf
EndFunc   ;==>OnChooseGPSBabelFolder

Func OnExitConfiguration()
	GUISetState(@SW_ENABLE, $g_oMain_GUI)
	GUIDelete($g_oConfig_GUI)
	;	GUISetState(@SW_HIDE, $g_oConfig_GUI)
	WinActivate($g_oMain_GUI)
EndFunc   ;==>OnExitConfiguration

Func OnSaveConfig()
	IniWrite(@ScriptDir & "\" & $g_sIniFilename, $g_sIniSection, $g_sIniKey_Raw, GUICtrlRead($g_oRawFolderField))
	IniWrite(@ScriptDir & "\" & $g_sIniFilename, $g_sIniSection, $g_sIniKey_XML, GUICtrlRead($g_oXmlFolderField))
	IniWrite(@ScriptDir & "\" & $g_sIniFilename, $g_sIniSection, $g_sIniKey_Excel, GUICtrlRead($g_oExcelFolderField))
	IniWrite(@ScriptDir & "\" & $g_sIniFilename, $g_sIniSection, $g_sIniKey_GPSBabel, GUICtrlRead($g_oGPSBabelFolderField))
	_ReadConfigurationFile()
EndFunc   ;==>OnSaveConfig

Func OnCancelConfig()
	OnExitConfiguration()
EndFunc   ;==>OnCancelConfig

Func OnFinishConfig()
	Local $l_iCompareRaw = StringCompare(GUICtrlRead($g_oRawFolderField), $g_sIniValue_RawFolder)
	Local $l_iCompareXML = StringCompare(GUICtrlRead($g_oXmlFolderField), $g_sIniValue_XMLFolder)
	Local $l_iCompareExcel = StringCompare(GUICtrlRead($g_oExcelFolderField), $g_sIniValue_ExcelFolder)

	Local $l_sMessageText1, $l_sMessageText2
	If $g_sChosenLanguage = "Norsk" Then
		$l_sMessageText1 = "Lagre endringer?"
		$l_sMessageText2 = "Du har endret hvilke foldere programmet leser fra og skriver til." & @CRLF & "Ønsker du å lagre endringene?"
	Else
		$l_sMessageText1 = "Store changes?"
		$l_sMessageText2 = "You have changed the folders this program reads from and writes to." & @CRLF & "Would you like to store these changes?"
	EndIf


	If (GUICtrlRead($g_oRawFolderField) <> $g_sIniValue_RawFolder Or GUICtrlRead($g_oXmlFolderField) <> $g_sIniValue_XMLFolder Or _
			GUICtrlRead($g_oExcelFolderField) <> $g_sIniValue_ExcelFolder Or GUICtrlRead($g_oGPSBabelFolderField) <> $g_sIniValue_GPSBabelFolder) Then
		If $IDYES = MsgBox($MB_YESNO, $l_sMessageText1, $l_sMessageText2) Then
			OnSaveConfig()
		EndIf
	EndIf

	OnExitConfiguration()
EndFunc   ;==>OnFinishConfig

Func OnChooseInputFile()
	; MsgBox($MB_SYSTEMMODAL, "You clicked on", "Velg en fil")

	Local $l_sFITdir
	If ($g_sIniValue_RawFolder = "") Then
		$l_sFITdir = @ScriptDir
	Else
		$l_sFITdir = $g_sIniValue_RawFolder
	EndIf

	; Display an open dialog to select a list of file(s).
	Local $l_sMessageText
	If $g_sChosenLanguage = "Norsk" Then
		$l_sMessageText = "Forstå meg! Velg filer fra Garmin-klokka"
	Else
		$l_sMessageText = "Understand how I am! Choose files from the Garmin watch"
	EndIf

	Local $l_sFileOpenDialog = FileOpenDialog($l_sMessageText, $l_sFITdir, "FIT files (*.fit)", $FD_FILEMUSTEXIST)
	If @error Then
		; Display the error message.
		; MsgBox($MB_SYSTEMMODAL, "", "No file(s) were selected.")

		; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
		;FileChangeDir(@ScriptDir)
	Else
		; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
		; FileChangeDir(@ScriptDir)

		; Replace instances of "|" with @CRLF in the string returned by FileOpenDialog.
		$l_sFileOpenDialog = StringReplace($l_sFileOpenDialog, "|", @CRLF)
		GUICtrlSetData($g_oInputField, $l_sFileOpenDialog)
		AddInfoToFileNameLabel()

		; Display the list of selected files.
		; MsgBox($MB_SYSTEMMODAL, "", "You chose the following files:" & @CRLF & $sFileOpenDialog)
	EndIf
EndFunc   ;==>OnChooseInputFile

; #FUNCTION _ControlFolderPaths# ====================================================================================================
; Name...........: _ControlFolderPaths
; Description....: Controlling that folders for selecting .fit files and storing xml and excel files are corect. And that the GPS-babel application is there.
;
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - Cannot find folder for .fit-files
;                  |2 - Cannot find folder for .xml-files
;                  |3 - Cannot find folder for .xlsx-files
;                  |4 - cannot find folder with GPSBabel application
;                  |5 - cannot find the application gpsbabel.exe with GPSBabel application
Func _ControlFolderPaths()


	If $g_sIniValue_RawFolder = "" Or FileExists($g_sIniValue_RawFolder) = 0 Then
		If $g_sChosenLanguage = "Norsk" Then
			MsgBox($MB_SYSTEMMODAL, "Finner ikke folderen", "Folderen som er definert for inputfiler fra Garmin-klokka er ikke gyldig.")
			Return SetError(1, @error, 0)
		Else
			MsgBox($MB_SYSTEMMODAL, "Can not find the folder", "The folder defined for input files from Garmin watch is not valid.")
			Return SetError(1, @error, 0)
		EndIf
	ElseIf $g_sIniValue_XMLFolder = "" Or FileExists($g_sIniValue_XMLFolder) = 0 Then
		If $g_sChosenLanguage = "Norsk" Then
			MsgBox($MB_SYSTEMMODAL, "Finner ikke folderen", "Folderen som er definert for å lagre Xml-filer er ikke gyldig.")
			Return SetError(2, @error, 0)
		Else
			MsgBox($MB_SYSTEMMODAL, "Can not find the folder", "The folder defined for storing XML-files is not valid.")
			Return SetError(2, @error, 0)
		EndIf
	ElseIf $g_sIniValue_ExcelFolder = "" Or FileExists($g_sIniValue_ExcelFolder) = 0 Then
		If $g_sChosenLanguage = "Norsk" Then
			MsgBox($MB_SYSTEMMODAL, "Finner ikke folderen", "Folderen som er definert for å lagre Excel-filer er ikke gyldig.")
			Return SetError(3, @error, 0)
		Else
			MsgBox($MB_SYSTEMMODAL, "Can not find the folder", "The folder defined for storing Excel-files is not valid.")
			Return SetError(3, @error, 0)
		EndIf
	ElseIf $g_sIniValue_GPSBabelFolder = "" Or FileExists($g_sIniValue_GPSBabelFolder) = 0 Then
		If $g_sChosenLanguage = "Norsk" Then
			MsgBox($MB_SYSTEMMODAL, "Finner ikke folderen", "Folderen som er definert for å finne applikasjonen GPSBabel er ikke gyldig.")
			Return SetError(4, @error, 0)
		Else
			MsgBox($MB_SYSTEMMODAL, "Can not find the folder", "The folder defined for finding the application GPSBabel is not valid.")
			Return SetError(4, @error, 0)
		EndIf
	ElseIf FileExists($g_sIniValue_GPSBabelFolder & "\gpsbabel.exe") = 0 Then
		If $g_sChosenLanguage = "Norsk" Then
			MsgBox($MB_SYSTEMMODAL, "Finner ikke applikasjonen", "Finner ikke applikasjonen gpsbabel.exe i angitt folder.")
			Return SetError(5, @error, 0)
		Else
			MsgBox($MB_SYSTEMMODAL, "Can not find the application", "Can not faind the application gpsbabel.exe in the given folder.")
			Return SetError(5, @error, 0)
		EndIf
	EndIf
	Return 1
EndFunc   ;==>_ControlFolderPaths


; #FUNCTION _TransformToXML# ====================================================================================================
; Name...........: _TransformToXML
; Description....: Making ready for transforming from FIT to XML.
;
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |6 - File to transfer is not chosen
;                  |7 - Cannot find the file with the given filename and path
;                  |8 - Output file already exists
;                  |9 - Error transforming to XML
Func _TransformToXML()

	Local $l_sInputFile, $l_sOutputFile, $l_sMessage1, $l_sMessage2
	;Local $l_sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
	$g_sCurrentFileName = GUICtrlRead($g_oOutputFileNameLabel) & GUICtrlRead($g_oOutputField)

	$l_sInputFile = GUICtrlRead($g_oInputField)
	$l_sOutputFile = $g_sIniValue_XMLFolder & "\" & $g_sCurrentFileName & ".xml"

	If $l_sInputFile = "" Then
		If $g_sChosenLanguage = "Norsk" Then
			$l_sMessage1 =  "Feil"
			$l_sMessage2 =  "Vennligst velg en fil som skal transformeres"
		Else
			$l_sMessage1 =  "Error"
			$l_sMessage2 =  "Please choose a file to be transformed"
		EndIf
		MsgBox($MB_SYSTEMMODAL, $l_sMessage1, $l_sMessage2)
		Return SetError(6, @error, 0)
	ElseIf FileExists($l_sInputFile) = 0 Then
		If $g_sChosenLanguage = "Norsk" Then
			$l_sMessage1 =  "Feil"
			$l_sMessage2 =  "Finner ikke filen. Vennligst velg en fil som skal transformeres"
		Else
			$l_sMessage1 =  "Error"
			$l_sMessage2 =  "Can not find the file. Please choose a file to be transformed"
		EndIf

		MsgBox($MB_SYSTEMMODAL, $l_sMessage1, $l_sMessage2)
		Return SetError(7, @error, 0)

	ElseIf _ControlOutputFileExistence() = 0 Then

		Return SetError(8, @error, 0)
	Else
		_TransformFile($l_sInputFile, $l_sOutputFile)
	EndIf
	Return 1
EndFunc   ;==>_TransformToXML

; #FUNCTION _ControlOutputFileExistence# ====================================================================================================
; Name...........: _ControlOutputFileExistence
; Description....: Checking if output files already exist. If either of them do, asks user to find a new file name for output file.
;
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |3 - XML file with chosen name already exists
;                  |4 - Exel file with chosen name already exists
Func _ControlOutputFileExistence()
	Local $l_sOutputXLSXFile, $l_sOutputXMLFile

	$l_sOutputXLSXFile = $g_sIniValue_ExcelFolder & "\" & $g_sCurrentFileName & ".xlsx"
	$l_sOutputXMLFile = $g_sIniValue_XMLFolder & "\" & $g_sCurrentFileName & ".xml"
	If FileExists($l_sOutputXMLFile) = 1 Then
		MsgBox($MB_SYSTEMMODAL, "Feil", "XML-filen eksisterer allerede. Velg et annet navn")
		Return SetError(3, @error, 0)
	ElseIf FileExists($l_sOutputXLSXFile) = 1 Then
		MsgBox($MB_SYSTEMMODAL, "Feil", "Excel-filen eksisterer allerede. Velg et annet navn")
		Return SetError(4, @error, 0)
	EndIf
	Return 1
EndFunc

Func _Format_Sheet_1_Descriptions()
	$g_oExcel.Application.ScreenUpdating = False ;Turn Screen Updating off
 	$g_oWorkbook.Sheets($g_sNameOfSheet_1_Descriptions).Columns($g_iS1ActivityDescriptionColumn).AutoFit
 	$g_oWorkbook.Sheets($g_sNameOfSheet_1_Descriptions).Columns($g_iS1EventDescriptionColumn).AutoFit


 	; Format main header area
 	With $g_oWorkbook.Sheets($g_sNameOfSheet_1_Descriptions).Range(_Excel_ColumnToLetter($g_iS1HeaderColumn) & $g_iS1HeaderRow & ":" & _Excel_ColumnToLetter($g_iS1HeaderColumn + 5) & $g_iS1HeaderRow)
 		.Merge
 		.Font.Bold = True
 		.Font.Color = $g_iTableMainHeaderFontColor
 		.Font.Size = $g_iTableMainHeaderFontSize
 		.Interior.Color = $g_iTableMainHeaderColorMaster
 	EndWith


	If $g_bS1ShowReduceVariationWarning = True Then
 		;Warning in yellow
 		With $g_oWorkbook.Sheets($g_sNameOfSheet_1_Descriptions).Range(_Excel_ColumnToLetter($g_iS1ReduceVariationWarningColumn) & $g_iS1ReduceVariationWarningRow & ":" & _Excel_ColumnToLetter($g_iS1ReduceVariationWarningColumn + 5) & $g_iS1ReduceVariationWarningRow)
 			.Merge ;Cells = True
 			.WrapText = True
 			.Font.Bold = True
 			.Interior.Color = $g_iPrivateColor_MEDIUM_YELLOW
 		EndWith
		;Example in yellow
 		With $g_oWorkbook.Sheets($g_sNameOfSheet_1_Descriptions).Range(_Excel_ColumnToLetter($g_iS1ReduceVariationExampleColumn ) & $g_iS1ReduceVariationExampleRow & ":" & _Excel_ColumnToLetter($g_iS1ReduceVariationExampleColumn + 5) & $g_iS1ReduceVariationExampleRow)
 			.Merge ;Cells = True
 			.WrapText = True
 			.Font.Bold = True
 			.Interior.Color = $g_iPrivateColor_MEDIUM_YELLOW
 		EndWith
 	EndIf

  	; Format activity area
  	With $g_oWorkbook.Sheets($g_sNameOfSheet_1_Descriptions).Range(_Excel_ColumnToLetter($g_iS1ActivityFirstColumn)&$g_iS1ActivityFirstHeaderRow&":"&_Excel_ColumnToLetter($g_iS1ActivityFirstColumn+ $g_iS1ActivityTable_NumberOfColumns-1)&$g_iS1ActivityFirstHeaderRow)
  		.Font.Bold = TRUE
  		.Font.Color = $g_iTableMainHeaderFontColor
  		.Font.Size = $g_iTableMainHeaderFontSize
  		.Interior.Color = $g_iTableMainHeaderColorActivity
  	EndWith
  	With $g_oWorkbook.Sheets($g_sNameOfSheet_1_Descriptions).Range(_Excel_ColumnToLetter($g_iS1ActivityFirstColumn)&$g_iS1ActivitySecondHeaderRow&":"&_Excel_ColumnToLetter($g_iS1ActivityFirstColumn+ $g_iS1ActivityTable_NumberOfColumns-1)&$g_iS1ActivitySecondHeaderRow)
  		.Font.Bold = FALSE
  		.Font.Color = $g_iTableSecondHeaderFontColor
  		.Interior.Color = $g_iTableSecondHeaderColorActivity
  	 EndWith

	Local $l_iCntr = $g_iS1ActivitySecondHeaderRow+1
  	While $l_iCntr < $g_iS1LastActivityContentRow +1 ; $g_iS1NumberOfActivities + 1 +$g_iS1ActivitySecondHeaderRow
 		$g_oWorkbook.Sheets($g_sNameOfSheet_1_Descriptions).Range(_Excel_ColumnToLetter($g_iS1ActivityFirstColumn)&$l_iCntr&":"&_Excel_ColumnToLetter($g_iS1ActivityFirstColumn+$g_iS1ActivityTable_NumberOfColumns-1)&$l_iCntr).Interior.Color = $g_iTableContentColorActivity
 		$l_iCntr = $l_iCntr + 2
 	WEnd

 	 ; Format event area
	With $g_oWorkbook.Sheets($g_sNameOfSheet_1_Descriptions).Range(_Excel_ColumnToLetter($g_iS1EventFirstColumn)&$g_iS1EventFirstHeaderRow&":"&_Excel_ColumnToLetter($g_iS1EventFirstColumn+ $g_iS1EventTable_NumberOfColumns-1)&$g_iS1EventFirstHeaderRow)
 		.Font.Bold = TRUE
 		.Font.Color = $g_iTableMainHeaderFontColor
 		.Font.Size = $g_iTableMainHeaderFontSize
 		.Interior.Color = $g_iTableMainHeaderColorEvent
 	EndWith
 	With $g_oWorkbook.Sheets($g_sNameOfSheet_1_Descriptions).Range(_Excel_ColumnToLetter($g_iS1EventFirstColumn)&$g_iS1EventSecondHeaderRow&":"&_Excel_ColumnToLetter($g_iS1EventFirstColumn+ $g_iS1EventTable_NumberOfColumns-1)&$g_iS1EventSecondHeaderRow)
 		.Font.Bold = FALSE
 		.Interior.Color = $g_iTableSecondHeaderColorEvent
 	EndWith
 	$l_iCntr = $g_iS1EventSecondHeaderRow +1
 	While $l_iCntr < $g_iS1LastEventContentRow +1 ;$g_iS1NumberOfEvents + 1 +$g_iS1EventSecondHeaderRow
 		$g_oWorkbook.Sheets($g_sNameOfSheet_1_Descriptions).Range(_Excel_ColumnToLetter($g_iS1EventFirstColumn)&$l_iCntr&":"&_Excel_ColumnToLetter($g_iS1EventFirstColumn+$g_iS1EventTable_NumberOfColumns-1)&$l_iCntr).Interior.Color = $g_iTableContentColorEvent
 		$l_iCntr = $l_iCntr + 2
 	WEnd

 	$g_oExcel.Application.ScreenUpdating = True ;Turn Screen Updating on
EndFunc   ;==>_Format_Sheet_1_Descriptions

Func _Fill_Sheet_1_Descriptions()

	$g_oExcel.Application.ScreenUpdating = False ;Turn Screen Updating off
	Local $l_sTextToWrite

	If $g_sChosenLanguage = "Norsk" Then
		$l_sTextToWrite = "Utvalgte aktiviteter og hendelser / atferder"
	Else
		$l_sTextToWrite = "Selected activities and events / behavior"
	EndIf

	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_1_Descriptions, $l_sTextToWrite, _Excel_ColumnToLetter($g_iS1HeaderColumn) & $g_iS1HeaderRow)


	;Info in yellow
	If $g_bS1ShowReduceVariationWarning = True then
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_1_Descriptions, $g_sS1ReduceVariationWarning, _Excel_ColumnToLetter($g_iS1ReduceVariationWarningColumn) & $g_iS1ReduceVariationWarningRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_1_Descriptions, $g_sS1ReduceVariationExample, _Excel_ColumnToLetter($g_iS1ReduceVariationExampleColumn) & $g_iS1ReduceVariationExampleRow)
	EndIf

	;Creating Activity Table
	If $g_sChosenLanguage = "Norsk" Then
		$l_sTextToWrite = "Aktivitet"
	Else
		$l_sTextToWrite = "Activity"
	EndIf
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_1_Descriptions, $l_sTextToWrite, _Excel_ColumnToLetter($g_iS1ActivityFirstColumn) & $g_iS1ActivityFirstHeaderRow)
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_1_Descriptions, "#", _Excel_ColumnToLetter($g_iS1ActivityFirstColumn) & $g_iS1ActivitySecondHeaderRow)

	If $g_sChosenLanguage = "Norsk" Then
		$l_sTextToWrite = "Beskrivelse"
	Else
		$l_sTextToWrite = "Description"
	EndIf
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_1_Descriptions, $l_sTextToWrite, _Excel_ColumnToLetter($g_iS1ActivityDescriptionColumn) & $g_iS1ActivitySecondHeaderRow)

	If $g_sChosenLanguage = "Norsk" Then
		$l_sTextToWrite = "[Aktivitet: Sett inn beskrivelse]"
	Else
		$l_sTextToWrite = "[Activities: Insert description]"
	EndIf
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_1_Descriptions, $l_sTextToWrite, _Excel_ColumnToLetter($g_iS1ActivityDescriptionColumn) & $g_iS1ActivityFirstContentRow & ":" & _Excel_ColumnToLetter($g_iS1ActivityDescriptionColumn) & $g_iS1LastActivityContentRow)
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_1_Descriptions, "1", _Excel_ColumnToLetter($g_iS1ActivityFirstColumn) & $g_iS1ActivityFirstContentRow)
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_1_Descriptions, "2", _Excel_ColumnToLetter($g_iS1ActivityFirstColumn) & $g_iS1ActivityFirstContentRow + 1)

	Local $l_sActivityRange = _Excel_ColumnToLetter($g_iS1ActivityFirstColumn) & $g_iS1ActivityFirstContentRow & ":" & _Excel_ColumnToLetter($g_iS1ActivityFirstColumn) & $g_iS1ActivityFirstContentRow + 1
	Local $l_sFillRange = _Excel_ColumnToLetter($g_iS1ActivityFirstColumn) & $g_iS1ActivityFirstContentRow & ":" & _Excel_ColumnToLetter($g_iS1ActivityFirstColumn) & $g_iS1LastActivityContentRow

	With $g_oWorkbook.Sheets($g_sNameOfSheet_1_Descriptions).Range($l_sActivityRange)
		.AutoFill($g_oWorkbook.Sheets($g_sNameOfSheet_1_Descriptions).Range($l_sFillRange), 0)
	EndWith



	;Creating Event Table
	If $g_sChosenLanguage = "Norsk" Then
		$l_sTextToWrite = "Hendelse / atferd"
	Else
		$l_sTextToWrite = "Events / Behavior"
	EndIf
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_1_Descriptions, $l_sTextToWrite, _Excel_ColumnToLetter($g_iS1EventFirstColumn) & $g_iS1EventFirstHeaderRow)
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_1_Descriptions, "#", _Excel_ColumnToLetter($g_iS1EventFirstColumn) & $g_iS1EventSecondHeaderRow)
	If $g_sChosenLanguage = "Norsk" Then
		$l_sTextToWrite = "Beskrivelse"
	Else
		$l_sTextToWrite = "Description"
	EndIf
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_1_Descriptions, $l_sTextToWrite, _Excel_ColumnToLetter($g_iS1EventDescriptionColumn) & $g_iS1EventSecondHeaderRow)
	If $g_sChosenLanguage = "Norsk" Then
		$l_sTextToWrite = "[Hendelse / atferd: Sett inn beskrivelse]"
	Else
		$l_sTextToWrite = "[Events / Behavior: Insert description]"
	EndIf
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_1_Descriptions, $l_sTextToWrite, _Excel_ColumnToLetter($g_iS1EventDescriptionColumn) & $g_iS1EventFirstContentRow & ":" & _Excel_ColumnToLetter($g_iS1EventDescriptionColumn) & $g_iS1LastEventContentRow)
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_1_Descriptions, "1", _Excel_ColumnToLetter($g_iS1EventFirstColumn) & $g_iS1EventFirstContentRow)
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_1_Descriptions, "2", _Excel_ColumnToLetter($g_iS1EventFirstColumn) & $g_iS1EventFirstContentRow + 1)

	Local $l_sEventRange = _Excel_ColumnToLetter($g_iS1EventFirstColumn) & $g_iS1EventFirstContentRow & ":" & _Excel_ColumnToLetter($g_iS1EventFirstColumn) & $g_iS1EventFirstContentRow + 1
	Local $l_sFillRange = _Excel_ColumnToLetter($g_iS1EventFirstColumn) & $g_iS1EventFirstContentRow & ":" & _Excel_ColumnToLetter($g_iS1EventFirstColumn) & $g_iS1LastEventContentRow

	With $g_oWorkbook.Sheets($g_sNameOfSheet_1_Descriptions).Range($l_sEventRange)
		.AutoFill($g_oWorkbook.Sheets($g_sNameOfSheet_1_Descriptions).Range($l_sFillRange), 0)
	EndWith

	$g_oExcel.Application.ScreenUpdating = True ;Turn Screen Updating on

EndFunc   ;==>_Fill_Sheet_1_Descriptions



Func _Format_Sheet_2_Observations()
	$g_oExcel.Application.ScreenUpdating = False ;Turn Screen Updating off
	$g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Columns($g_iS2ActivityDescriptionColumn).AutoFit
	$g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Columns($g_iS2EventDescriptionColumn).AutoFit
	$g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Columns($g_iS2EventToColumn).AutoFit

	; Format main header area
	With $g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range(_Excel_ColumnToLetter($g_iS2TimeSummaryHeaderColumn) & $g_iS2TimeSummaryHeaderRow & ":" & _Excel_ColumnToLetter($g_iS2TimeSummaryHeaderColumn + 1) & $g_iS2TimeSummaryHeaderRow)
		.Merge
	EndWith
	With $g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range(_Excel_ColumnToLetter($g_iS2HourCorrectionInstructionColumn) & $g_iS2HourCorrectionRow & ":" & _Excel_ColumnToLetter($g_iS2HourCorrectionInputColumn - 1) & $g_iS2HourCorrectionRow)
		.Merge
	EndWith
	With $g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range(_Excel_ColumnToLetter($g_iS2TimeSummaryHeaderColumn) & $g_iS2TimeSummaryHeaderRow & ":" & _Excel_ColumnToLetter($g_iS2TimeSummaryToColumn) & $g_iS2TimeSummaryHeaderRow)
		.Font.Bold = True
		.Font.Color = $g_iTableMainHeaderFontColor
		.Font.Size = $g_iTableMainHeaderFontSize
		.Interior.Color = $g_iTableMainHeaderColorMaster
	EndWith
	With $g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range(_Excel_ColumnToLetter($g_iS2TimeSummaryHeaderColumn) & $g_iS2TimeSummaryContentRow & ":" & _Excel_ColumnToLetter($g_iS2TimeSummaryFromColumn - 1) & $g_iS2TimeSummaryContentRow)
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableMainHeaderColorMaster
	EndWith
	With $g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range(_Excel_ColumnToLetter($g_iS2HourCorrectionInstructionColumn) & $g_iS2HourCorrectionRow & ":" & _Excel_ColumnToLetter($g_iS2HourCorrectionInputColumn - 1) & $g_iS2HourCorrectionRow)
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableMainHeaderColorMaster
	EndWith
	With $g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range(_Excel_ColumnToLetter($g_iS2TimeSummaryFromColumn) & $g_iS2TimeSummaryContentRow & ":" & _Excel_ColumnToLetter($g_iS2TimeSummaryToColumn) & $g_iS2TimeSummaryContentRow)
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableSecondHeaderColorMaster
	EndWith

	With $g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range(_Excel_ColumnToLetter($g_iS2TimeSummaryHeaderColumn) & $g_iS2TimeSummaryContentRow & ":" & _Excel_ColumnToLetter($g_iS2TimeSummaryToColumn) & $g_iS2HourCorrectionRow).Borders
		.Linestyle = $g_xlContinuous
		.Color = $g_iTableMainHeaderColorMaster
		.Weight = $g_iTableInsideBorderThickness

	EndWith

	With $g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range(_Excel_ColumnToLetter($g_iS2TimeSummaryHeaderColumn) & $g_iS2TimeSummaryHeaderRow & ":" & _Excel_ColumnToLetter($g_iS2TimeSummaryToColumn) & $g_iS2HourCorrectionRow)
		.Borders($g_xlEdgeBottom).Weight = $g_iTableOutsideBorderThickness
		.Borders($g_xlEdgeTop).Weight = $g_iTableOutsideBorderThickness
		.Borders($g_xlEdgeLeft).Weight = $g_iTableOutsideBorderThickness
		.Borders($g_xlEdgeRight).Weight = $g_iTableOutsideBorderThickness
	EndWith

	If $g_bS2ShowReduceVariationWarning = True Then
		;Warning in yellow
		With $g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range(_Excel_ColumnToLetter($g_iS2ReduceVariationWarningColumn) & $g_iS2ReduceVariationWarningRow & ":" & _Excel_ColumnToLetter($g_iS2ReduceVariationWarningColumn + 4) & $g_iS2ReduceVariationWarningRow)
			.Merge ;Cells = True
			.WrapText = True
			.Font.Bold = True
			.Interior.Color = $g_iPrivateColor_MEDIUM_YELLOW
		EndWith
	EndIf

 	 ; Format activity area
 	 With $g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range(_Excel_ColumnToLetter($g_iS2ActivityFirstColumn)&$g_iS2ActivityFirstHeaderRow&":"&_Excel_ColumnToLetter($g_iS2ActivityFirstColumn+ $g_iS2ActivityTable_NumberOfColumns-1)&$g_iS2ActivityFirstHeaderRow)
 		.Font.Bold = TRUE
 		.Font.Color = $g_iTableMainHeaderFontColor
 		.Font.Size = $g_iTableMainHeaderFontSize
 		.Interior.Color = $g_iTableMainHeaderColorActivity
 	EndWith
 	With $g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range(_Excel_ColumnToLetter($g_iS2ActivityFirstColumn)&$g_iS2ActivitySecondHeaderRow&":"&_Excel_ColumnToLetter($g_iS2ActivityFirstColumn+ $g_iS2ActivityTable_NumberOfColumns-1)&$g_iS2ActivitySecondHeaderRow)
 		.Font.Bold = FALSE
 		.Font.Color = $g_iTableSecondHeaderFontColor
 		.Interior.Color = $g_iTableSecondHeaderColorActivity
 	 EndWith

 	 Local $l_iCntr = $g_iS2ActivitySecondHeaderRow+1
 	While $l_iCntr < $g_iNumberOfActivities + 1 +$g_iS2ActivitySecondHeaderRow
 		$g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range(_Excel_ColumnToLetter($g_iS2ActivityFirstColumn)&$l_iCntr&":"&_Excel_ColumnToLetter($g_iS2ActivityFirstColumn+$g_iS2ActivityTable_NumberOfColumns-1)&$l_iCntr).Interior.Color = $g_iTableContentColorActivity
 		$l_iCntr = $l_iCntr + 2
 	WEnd

 	 ; Format event area
 	 With $g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range(_Excel_ColumnToLetter($g_iS2EventFirstColumn)&$g_iS2EventFirstHeaderRow&":"&_Excel_ColumnToLetter($g_iS2EventFirstColumn+ $g_iS2EventTable_NumberOfColumns-1)&$g_iS2EventFirstHeaderRow)
 		.Font.Bold = TRUE
 		.Font.Color = $g_iTableMainHeaderFontColor
 		.Font.Size = $g_iTableMainHeaderFontSize
 		.Interior.Color = $g_iTableMainHeaderColorEvent
 	EndWith
 	With $g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range(_Excel_ColumnToLetter($g_iS2EventFirstColumn)&$g_iS2EventSecondHeaderRow&":"&_Excel_ColumnToLetter($g_iS2EventFirstColumn+ $g_iS2EventTable_NumberOfColumns-1)&$g_iS2EventSecondHeaderRow)
 		.Font.Bold = FALSE
 		.Interior.Color = $g_iTableSecondHeaderColorEvent
 	EndWith
 	$l_iCntr = $g_iS2EventSecondHeaderRow +1
 	While $l_iCntr < $g_iNumberOfEvents + 1 +$g_iS2EventSecondHeaderRow
 		$g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range(_Excel_ColumnToLetter($g_iS2EventFirstColumn)&$l_iCntr&":"&_Excel_ColumnToLetter($g_iS2EventFirstColumn+$g_iS2EventTable_NumberOfColumns-1)&$l_iCntr).Interior.Color = $g_iTableContentColorEvent
 		$l_iCntr = $l_iCntr + 2
 	WEnd

	$g_oExcel.Application.ScreenUpdating = True ;Turn Screen Updating on
EndFunc   ;==>_Format_Sheet_2_Observations
Func _Fill_Sheet_2_Observations($l_iNumberOfTrackpoints)
	;$g_sNameOfSheet_2_Observations: Inserting time range header
	$g_oExcel.Application.ScreenUpdating = False ;Turn Screen Updating off
	If $g_sChosenLanguage = "Norsk" Then
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "Måleperiode", _Excel_ColumnToLetter($g_iS2TimeSummaryHeaderColumn) & $g_iS2TimeSummaryHeaderRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "Fra kl", _Excel_ColumnToLetter($g_iS2TimeSummaryFromColumn) & $g_iS2TimeSummaryHeaderRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "Til kl", _Excel_ColumnToLetter($g_iS2TimeSummaryToColumn) & $g_iS2TimeSummaryHeaderRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "Er måleperioden feil?  Legg til eller trekk fra timer her:", _Excel_ColumnToLetter($g_iS2HourCorrectionInstructionColumn) & $g_iS2HourCorrectionRow)
	Else
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "Recorded period", _Excel_ColumnToLetter($g_iS2TimeSummaryHeaderColumn) & $g_iS2TimeSummaryHeaderRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "From", _Excel_ColumnToLetter($g_iS2TimeSummaryFromColumn) & $g_iS2TimeSummaryHeaderRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "To", _Excel_ColumnToLetter($g_iS2TimeSummaryToColumn) & $g_iS2TimeSummaryHeaderRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "Is the recorded period wrong? Add or subtract an hour:", _Excel_ColumnToLetter($g_iS2HourCorrectionInstructionColumn) & $g_iS2HourCorrectionRow)
	EndIf
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, $g_iS2HourCorrectionValue, _Excel_ColumnToLetter($g_iS2HourCorrectionInputColumn) & $g_iS2HourCorrectionRow)
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "00:00:00", _Excel_ColumnToLetter($g_iS2TimeSummaryFromColumn) & $g_iS2TimeSummaryContentRow & ":" & _Excel_ColumnToLetter($g_iS2TimeSummaryToColumn) & $g_iS2TimeSummaryContentRow)
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "='"& $g_sNameOfSheet_5_CalculationsMainGraph&"'!"&_Excel_ColumnToLetter($g_iS5Column_CorrectedTime)&$g_iS5FirstContentRow , _Excel_ColumnToLetter($g_iS2TimeSummaryFromColumn) & $g_iS2TimeSummaryContentRow)
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "='"& $g_sNameOfSheet_5_CalculationsMainGraph&"'!"&_Excel_ColumnToLetter($g_iS5Column_CorrectedTime)&$g_iS5LastContentRow , _Excel_ColumnToLetter($g_iS2TimeSummaryToColumn) & $g_iS2TimeSummaryContentRow)



	;Warning in yellow
	If $g_bS2ShowReduceVariationWarning = True then
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, $g_sS2ReduceVariationWarning, _Excel_ColumnToLetter($g_iS2ReduceVariationWarningColumn) & $g_iS2ReduceVariationWarningRow)
	EndIf


	;Creating Activity Table
	Local $l_sActivityRange = _Excel_ColumnToLetter($g_iS2ActivityFirstColumn) & $g_iS2ActivitySecondHeaderRow & ":" & _Excel_ColumnToLetter($g_iS2ActivityToColumn) & $g_iS2LastActivityContentRow-1
	Local $l_xlSrcRange = 1

	If $g_sChosenLanguage = "Norsk" Then
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "Aktivitet", _Excel_ColumnToLetter($g_iS2ActivityFirstColumn) & $g_iS2ActivityFirstHeaderRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "Beskrivelse", _Excel_ColumnToLetter($g_iS2ActivityDescriptionColumn) & $g_iS2ActivitySecondHeaderRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "Fra klokken", _Excel_ColumnToLetter($g_iS2ActivityFromColumn) & $g_iS2ActivitySecondHeaderRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "Til klokken", _Excel_ColumnToLetter($g_iS2ActivityFirstColumn + 3) & $g_iS2ActivitySecondHeaderRow)

		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "[Aktivitet: Velg fra nedtrekksmeny eller skriv]", _Excel_ColumnToLetter($g_iS2ActivityDescriptionColumn) & $g_iS2ActivityFirstContentRow & ":" & _Excel_ColumnToLetter($g_iS2ActivityDescriptionColumn) & $g_iS2LastActivityContentRow)
	Else
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "Activity", _Excel_ColumnToLetter($g_iS2ActivityFirstColumn) & $g_iS2ActivityFirstHeaderRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "Description", _Excel_ColumnToLetter($g_iS2ActivityDescriptionColumn) & $g_iS2ActivitySecondHeaderRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "From", _Excel_ColumnToLetter($g_iS2ActivityFromColumn) & $g_iS2ActivitySecondHeaderRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "To", _Excel_ColumnToLetter($g_iS2ActivityFirstColumn + 3) & $g_iS2ActivitySecondHeaderRow)

		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "[Activity: Choose from dropdown menu or write]", _Excel_ColumnToLetter($g_iS2ActivityDescriptionColumn) & $g_iS2ActivityFirstContentRow & ":" & _Excel_ColumnToLetter($g_iS2ActivityDescriptionColumn) & $g_iS2LastActivityContentRow)
	EndIf
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "#", _Excel_ColumnToLetter($g_iS2ActivityFirstColumn) & $g_iS2ActivitySecondHeaderRow)
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "1", _Excel_ColumnToLetter($g_iS2ActivityFirstColumn) & $g_iS2ActivityFirstContentRow)
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "2", _Excel_ColumnToLetter($g_iS2ActivityFirstColumn) & $g_iS2ActivityFirstContentRow + 1)

	$l_sActivityRange = _Excel_ColumnToLetter($g_iS2ActivityFirstColumn) & $g_iS2ActivityFirstContentRow & ":" & _Excel_ColumnToLetter($g_iS2ActivityFirstColumn) & $g_iS2ActivityFirstContentRow + 1
	Local $l_sFillRange = _Excel_ColumnToLetter($g_iS2ActivityFirstColumn) & $g_iS2ActivityFirstContentRow & ":" & _Excel_ColumnToLetter($g_iS2ActivityFirstColumn) & $g_iS2LastActivityContentRow

	With $g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range($l_sActivityRange)
		.AutoFill($g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range($l_sFillRange), 0)
	EndWith

	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "00:00:00", _Excel_ColumnToLetter($g_iS2ActivityFromColumn) & $g_iS2ActivityFirstContentRow & ":" & _Excel_ColumnToLetter($g_iS2ActivityToColumn) & $g_iS2LastActivityContentRow)

	;Add data validation for activity table
	$l_sActivityRange = _Excel_ColumnToLetter($g_iS2ActivityDescriptionColumn) & $g_iS2ActivityFirstContentRow & ":" & _Excel_ColumnToLetter($g_iS2ActivityDescriptionColumn) & $g_iS2LastActivityContentRow
	Local $l_sValidationRange = "='"&$g_sNameOfSheet_1_Descriptions &"'!$"&_Excel_ColumnToLetter($g_iS1ActivityDescriptionColumn)&"$"&$g_iS1ActivityFirstContentRow & ":$" & _Excel_ColumnToLetter($g_iS1ActivityDescriptionColumn) &"$"& $g_iS1LastActivityContentRow
	_Excel_RangeValidate($g_oWorkbook, $g_sNameOfSheet_2_Observations, $l_sActivityRange, $xlValidateList, $l_sValidationRange, Default, Default, True,Default, "", "")

	With $g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range($l_sActivityRange)
		.Validation.ShowError = False
	EndWith

	;Creating Event Table
	Local $l_sEventRange = _Excel_ColumnToLetter($g_iS2EventFirstColumn) & $g_iS2EventSecondHeaderRow & ":" & _Excel_ColumnToLetter($g_iS2EventToColumn) & $g_iS2LastEventContentRow-1

	If $g_sChosenLanguage = "Norsk" Then
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "Hendelse / atferd", _Excel_ColumnToLetter($g_iS2EventFirstColumn) & $g_iS2EventFirstHeaderRow)

		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "Beskrivelse", _Excel_ColumnToLetter($g_iS2EventDescriptionColumn) & $g_iS2EventSecondHeaderRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "Fra klokken", _Excel_ColumnToLetter($g_iS2EventFromColumn) & $g_iS2EventSecondHeaderRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "Tidspunkt Slutt (Valgfritt å endre)", _Excel_ColumnToLetter($g_iS2EventToColumn) & $g_iS2EventSecondHeaderRow)

		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "[Hendelse/atferd: Velg fra nedtrekksmeny eller skriv]", _Excel_ColumnToLetter($g_iS2EventDescriptionColumn) & $g_iS2EventFirstContentRow & ":" & _Excel_ColumnToLetter($g_iS2EventDescriptionColumn) & $g_iS2LastEventContentRow)
	Else
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "Events / Behavior", _Excel_ColumnToLetter($g_iS2EventFirstColumn) & $g_iS2EventFirstHeaderRow)

		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "Description", _Excel_ColumnToLetter($g_iS2EventDescriptionColumn) & $g_iS2EventSecondHeaderRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "From", _Excel_ColumnToLetter($g_iS2EventFromColumn) & $g_iS2EventSecondHeaderRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "To  (Optional to change)", _Excel_ColumnToLetter($g_iS2EventToColumn) & $g_iS2EventSecondHeaderRow)

		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "[Events / Behavior: Choose from dropdown menu or write]", _Excel_ColumnToLetter($g_iS2EventDescriptionColumn) & $g_iS2EventFirstContentRow & ":" & _Excel_ColumnToLetter($g_iS2EventDescriptionColumn) & $g_iS2LastEventContentRow)

	EndIf
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "#", _Excel_ColumnToLetter($g_iS2EventFirstColumn) & $g_iS2EventSecondHeaderRow)
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "1", _Excel_ColumnToLetter($g_iS2EventFirstColumn) & $g_iS2EventFirstContentRow)
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "2", _Excel_ColumnToLetter($g_iS2EventFirstColumn) & $g_iS2EventFirstContentRow + 1)

	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "00:00:00", _Excel_ColumnToLetter($g_iS2EventFromColumn) & $g_iS2EventFirstContentRow & ":" & _Excel_ColumnToLetter($g_iS2EventToColumn) & $g_iS2LastEventContentRow)
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "=" & _Excel_ColumnToLetter($g_iS2EventFromColumn) & $g_iS2EventFirstContentRow, _Excel_ColumnToLetter($g_iS2EventToColumn) & $g_iS2EventFirstContentRow)
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_2_Observations, "=" & _Excel_ColumnToLetter($g_iS2EventFromColumn) & $g_iS2EventFirstContentRow+1, _Excel_ColumnToLetter($g_iS2EventToColumn) & $g_iS2EventFirstContentRow+1)

	$l_sEventRange = _Excel_ColumnToLetter($g_iS2EventFirstColumn) & $g_iS2EventFirstContentRow & ":" & _Excel_ColumnToLetter($g_iS2EventToColumn) & $g_iS2EventFirstContentRow + 1
	$l_sFillRange = _Excel_ColumnToLetter($g_iS2EventFirstColumn) & $g_iS2EventFirstContentRow & ":" & _Excel_ColumnToLetter($g_iS2EventToColumn) & $g_iS2LastEventContentRow

	With $g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range($l_sEventRange)
		.AutoFill($g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range($l_sFillRange), 0)
	EndWith

	;Add data validation for activity table
	$l_sEventRange = _Excel_ColumnToLetter($g_iS2EventDescriptionColumn) & $g_iS2EventFirstContentRow & ":" & _Excel_ColumnToLetter($g_iS2EventDescriptionColumn) & $g_iS2LastEventContentRow
	$l_sValidationRange = "='"&$g_sNameOfSheet_1_Descriptions &"'!$"&_Excel_ColumnToLetter($g_iS1EventDescriptionColumn)&"$"&$g_iS1EventFirstContentRow & ":$" & _Excel_ColumnToLetter($g_iS1EventDescriptionColumn) &"$"& $g_iS1LastEventContentRow
	_Excel_RangeValidate($g_oWorkbook, $g_sNameOfSheet_2_Observations, $l_sEventRange, $xlValidateList, $l_sValidationRange,Default,Default, False,Default, "", "")
	With $g_oWorkbook.Sheets($g_sNameOfSheet_2_Observations).Range($l_sEventRange)
		.Validation.ShowError = False
	EndWith

	$g_oExcel.Application.ScreenUpdating = True ;Turn Screen Updating on

EndFunc   ;==>_Fill_Sheet_2_Observations

Func _Format_Sheet_3_MainGraph()
	$g_oExcel.Application.ScreenUpdating = False ;Turn Screen Updating off
	;	Time symmary
	;Header
	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3TimeSummaryHeaderColumn) & $g_iS3TimeSummaryHeaderRow & ":" & _Excel_ColumnToLetter($g_iS3TimeSummaryToColumn) & $g_iS3TimeSummaryHeaderRow)
		.Font.Bold = True
		.Font.Color = $g_iTableMainHeaderFontColor
		.Font.Size = $g_iTableMainHeaderFontSize
		.Interior.Color = $g_iTableMainHeaderColorMaster
	EndWith
	;All introtext in first columns
	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3TimeSummaryHeaderColumn) & $g_iS3TimeSummaryContentRow & ":" & _Excel_ColumnToLetter($g_iS3TimeSummaryFromColumn - 1) & $g_iS3GraphNumberOfDatapointsRow)
		;.MergeCells = True
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableMainHeaderColorMaster
	EndWith
	;Introtext for data resolution
	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3GraphTimeResolutionIntroColumn) & $g_iS3GraphTimeResolutionRow & ":" & _Excel_ColumnToLetter($g_iS3TimeSummaryFromColumn - 1) & $g_iS3GraphTimeResolutionRow)
		.Merge;Cells = True
		.WrapText = True
		.ShrinkToFit = False
	EndWith
	;Time summary contents
	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3TimeSummaryFromColumn) & $g_iS3TimeSummaryContentRow & ":" & _Excel_ColumnToLetter($g_iS3TimeSummaryToColumn) & $g_iS3TimeSummaryContentRow)
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableSecondHeaderColorMaster
	EndWith

	;Unused To-column-cells
	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3TimeSummaryToColumn) & $g_iS3GraphTimeResolutionRow & ":" & _Excel_ColumnToLetter($g_iS3TimeSummaryToColumn) & $g_iS3GraphNumberOfDatapointsRow )
		;.MergeCells = True
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableMainHeaderColorMaster
	EndWith

	;Number of datapoints
	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3GraphNumberOfDatapointsValueColumn) & $g_iS3GraphNumberOfDatapointsRow)
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableSecondHeaderColorMaster
	EndWith

	;Borders around
	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3TimeSummaryHeaderColumn) & $g_iS3TimeSummaryHeaderRow & ":" & _Excel_ColumnToLetter($g_iS3TimeSummaryToColumn) & $g_iS3GraphTimeResolutionRow)
		;.WrapText = True
		.Borders.LineStyle = $g_xlContinuous
		.Borders.Weight = $xlThick
		.Borders.Color = $g_iTableMainHeaderColorMaster
	EndWith

	If $g_bS3ShowResolutionWarning = True Then
		;Warning in yellow
		With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3GraphTimeResolutionWarningColumn) & $g_iS3GraphTimeResolutionRow & ":" & _Excel_ColumnToLetter($g_iS3GraphTimeResolutionWarningColumn + 5) & $g_iS3GraphTimeResolutionRow)
			.Merge ;Cells = True
			.WrapText = True
			.Font.Bold = True
			.Interior.Color = $g_iPrivateColor_MEDIUM_YELLOW
		EndWith
	EndIf

	With $g_oWorkbook.Worksheets($g_sNameOfSheet_3_GraphMain).Rows($g_iS3GraphTimeResolutionRow)
        .RowHeight = 30
    EndWith

	;Format Pulse summary
	;Header
	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3PulseSummaryIntroColumn) & $g_iS3PulseSummaryHeaderRow & ":" & _Excel_ColumnToLetter($g_iS3TimeSummaryToColumn) & $g_iS3PulseSummaryHeaderRow)
		.Font.Bold = True
		.Font.Color = $g_iTableMainHeaderFontColor
		.Font.Size = $g_iTableMainHeaderFontSize
		.Interior.Color = $g_iTableMainHeaderColorMaster
	EndWith
	;All introtext in first columns
	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3PulseSummaryIntroColumn) & $g_iS3PulseSummaryAverageRow & ":" & _Excel_ColumnToLetter($g_iS3TimeSummaryFromColumn) & $g_iS3PulseSummaryAvPlusTwoStDevRow)
		;.MergeCells = True
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableMainHeaderColorMaster
	EndWith

	;Pulse summary Values
	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3PulseSummaryValueColumn) & $g_iS3PulseSummaryAverageRow & ":" & _Excel_ColumnToLetter($g_iS3PulseSummaryValueColumn) & $g_iS3PulseSummaryAvPlusTwoStDevRow)
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableSecondHeaderColorMaster
		.NumberFormat = "0,00"
	EndWith

	;Borders around
	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3PulseSummaryIntroColumn) & $g_iS3PulseSummaryHeaderRow & ":" & _Excel_ColumnToLetter($g_iS3PulseSummaryValueColumn) & $g_iS3PulseSummaryAvPlusTwoStDevRow)
		;.WrapText = True
		.Borders.LineStyle = $g_xlContinuous
		.Borders.Weight = $xlThick
		.Borders.Color = $g_iTableMainHeaderColorMaster
	EndWith


	;Format Activity
	;	Green header
	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3ActivityIntroColumn) & $g_iS3ActivityHeaderRow & ":" & _Excel_ColumnToLetter($g_iS3ActivityValueColumn) & $g_iS3ActivityHeaderRow)
		.MergeCells = True
		.Font.Bold = True
		.Font.Color = $g_iTableMainHeaderFontColor
		.Font.Size = $g_iTableMainHeaderFontSize
		.Interior.Color = $g_iTableMainHeaderColorActivity
	EndWith
	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3ActivityIntroColumn) & $g_iS3ActivityLevelRow & ":" & _Excel_ColumnToLetter($g_iS3ActivityValueColumn - 1) & $g_iS3ActivityLevelRow)
		.MergeCells = True
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableMainHeaderColorActivity
	EndWith
	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3ActivityIntroColumn) & $g_iS3ActivityDistanceRow & ":" & _Excel_ColumnToLetter($g_iS3ActivityValueColumn - 1) & $g_iS3ActivityDistanceRow)
		.MergeCells = True
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableMainHeaderColorActivity
	EndWith
	;Borders around input cells
	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3ActivityValueColumn) & $g_iS3ActivityLevelRow & ":" & _Excel_ColumnToLetter($g_iS3ActivityValueColumn) & $g_iS3ActivityDistanceRow )
		.WrapText = True
		.Borders.LineStyle = $g_xlContinuous
		.Borders.Weight = $xlThick
		.Borders.Color = $g_iTableMainHeaderColorActivity
	EndWith

	;Format Event
	;	Yellow header
	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3EventIntroColumn) & $g_iS3EventHeaderRow & ":" & _Excel_ColumnToLetter($g_iS3EventValueColumn) & $g_iS3EventHeaderRow)
		.MergeCells = True
		.Font.Bold = True
		.Font.Color = $g_iTableMainHeaderFontColor
		.Font.Size = $g_iTableMainHeaderFontSize
		.Interior.Color = $g_iTableMainHeaderColorEvent
	EndWith

	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3EventIntroColumn) & $g_iS3EventLevelRow & ":" & _Excel_ColumnToLetter($g_iS3EventValueColumn - 1) & $g_iS3EventLevelRow)
		.MergeCells = True
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableMainHeaderColorEvent
	EndWith
	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3EventIntroColumn) & $g_iS3EventDistanceRow & ":" & _Excel_ColumnToLetter($g_iS3EventValueColumn - 1) & $g_iS3EventDistanceRow)
		.MergeCells = True
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableMainHeaderColorEvent
	EndWith
	;Borders around input cells
	With $g_oWorkbook.Sheets($g_sNameOfSheet_3_GraphMain).Range(_Excel_ColumnToLetter($g_iS3EventValueColumn) & $g_iS3EventLevelRow & ":" & _Excel_ColumnToLetter($g_iS3EventValueColumn) & $g_iS3EventDistanceRow )
		.WrapText = True
		.Borders.LineStyle = $g_xlContinuous
		.Borders.Weight = $xlThick
		.Borders.Color = $g_iTableMainHeaderColorEvent
	EndWith
	$g_oExcel.Application.ScreenUpdating = True ;Turn Screen Updating on
EndFunc   ;==>_Format_Sheet_3_MainGraph


Func _Fill_Sheet_3_MainGraph()
	$g_oExcel.Application.ScreenUpdating = False ;Turn Screen Updating off
	;$g_sNameOfSheet_2_Observations: Inserting time range header

	Local $l_sCreateTimeFormula = "=ROUNDDOWN(COUNT('" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!"& _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5FirstContentRow & ":" & _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5LastSecondIntervalRow&")/$"& _Excel_ColumnToLetter($g_iS3GraphNumberOfDatapointsValueColumn) & "$" & $g_iS3GraphTimeResolutionRow &";0)"


	Local $l_sLongPulseRange = "'"&$g_sNameOfSheet_5_CalculationsMainGraph&"'!$"& _Excel_ColumnToLetter($g_iS5Column_PulseForLongGraph) & "$"& $g_iS5FirstContentRow &":$" & _Excel_ColumnToLetter($g_iS5Column_PulseForLongGraph) & "$"& $g_iS5LastSecondIntervalRow
	Local $l_sFormula_Average =  "=AVERAGEIF("&$l_sLongPulseRange  & ";"& Chr(34) &"<>#N/A"&Chr(34)&")"
	Local $l_sFormula_StDev = "=AGGREGATE( 8;6;"&$l_sLongPulseRange  & ")"
	Local $l_sFormula_AvPlusOneStDev = "=AGGREGATE( 8;6;"&$l_sLongPulseRange  & ") + " &_Excel_ColumnToLetter($g_iS3PulseSummaryValueColumn) & $g_iS3PulseSummaryAverageRow
	Local $l_sFormula_AvPlusTwoStDev = "=2*AGGREGATE( 8;6;"&$l_sLongPulseRange  & ") + " &_Excel_ColumnToLetter($g_iS3PulseSummaryValueColumn) & $g_iS3PulseSummaryAverageRow
;	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_3_GraphMain, $l_sFormula, _Excel_ColumnToLetter($g_iS3PulseSummaryValueColumn) & $g_iS3PulseSummaryAverageRow)


	Local $l_iCntr = 0
	Local $l_iGraphSheetTableWidth = $g_iS3GraphTimeResolutionWarningColumn
	Local $l_iGraphSheetTableHeight = $g_iS3ActivityDistanceRow + 3
	Local $l_aGraphSheetArray[$l_iGraphSheetTableHeight][$l_iGraphSheetTableWidth]

	If $g_sChosenLanguage = "Norsk" Then
		$l_aGraphSheetArray[$g_iS3TimeSummaryHeaderRow-1][$g_iS3TimeSummaryHeaderColumn-1] = "Måleperiode"
		$l_aGraphSheetArray[$g_iS3TimeSummaryHeaderRow-1][$g_iS3TimeSummaryFromColumn-1] = "Fra kl"
		$l_aGraphSheetArray[$g_iS3TimeSummaryHeaderRow-1][$g_iS3TimeSummaryToColumn-1] = "Til kl"
		$l_aGraphSheetArray[$g_iS3GraphTimeResolutionRow-1][$g_iS3GraphTimeResolutionIntroColumn-1] = "Velg oppløsning for hovedgrafen:"

		$l_aGraphSheetArray[$g_iS3GraphTimeResolutionRow-1][$g_iS3GraphTimeResolutionInputColumn-1] = "5"
		$l_aGraphSheetArray[$g_iS3GraphTimeResolutionRow-1][$g_iS3TimeSummaryToColumn-1] = " sekunder"
		$l_aGraphSheetArray[$g_iS3GraphNumberOfDatapointsRow-1][$g_iS3GraphNumberOfDatapointsIntroColumn-1] = "Tilsvarer"
		$l_aGraphSheetArray[$g_iS3GraphNumberOfDatapointsRow-1][$g_iS3GraphNumberOfDatapointsValueColumn-1] = $l_sCreateTimeFormula
		$l_aGraphSheetArray[$g_iS3GraphNumberOfDatapointsRow-1][$g_iS3GraphNumberOfDatapointsText2Column-1] = "datapunkter"

		$l_aGraphSheetArray[$g_iS3PulseSummaryHeaderRow-1][$g_iS3PulseSummaryIntroColumn-1] = "Oppsummering puls "
		$l_aGraphSheetArray[$g_iS3PulseSummaryAverageRow-1][$g_iS3PulseSummaryIntroColumn-1] = "Gjennomsnitt"
		$l_aGraphSheetArray[$g_iS3PulseSummaryAverageRow-1][$g_iS3PulseSummaryValueColumn-1] = $l_sFormula_Average
		$l_aGraphSheetArray[$g_iS3PulseSummaryStDevRow-1][$g_iS3PulseSummaryIntroColumn-1] = "Standardavvik"
		$l_aGraphSheetArray[$g_iS3PulseSummaryStDevRow-1][$g_iS3PulseSummaryValueColumn-1] = $l_sFormula_StDev
		$l_aGraphSheetArray[$g_iS3PulseSummaryAvPlusOneStDevRow-1][$g_iS3PulseSummaryIntroColumn-1] = "Gjennomsnitt pluss ett standardavvik"
		$l_aGraphSheetArray[$g_iS3PulseSummaryAvPlusOneStDevRow-1][$g_iS3PulseSummaryValueColumn-1] = $l_sFormula_AvPlusOneStDev
		$l_aGraphSheetArray[$g_iS3PulseSummaryAvPlusTwoStDevRow-1][$g_iS3PulseSummaryIntroColumn-1] ="Gjennomsnitt pluss to standardavvik"
		$l_aGraphSheetArray[$g_iS3PulseSummaryAvPlusTwoStDevRow-1][$g_iS3PulseSummaryValueColumn-1] = $l_sFormula_AvPlusTwoStDev

		$l_aGraphSheetArray[$g_iS3EventHeaderRow-1][$g_iS3EventIntroColumn-1] = "Hendelser"
		$l_aGraphSheetArray[$g_iS3EventLevelRow -1][$g_iS3EventIntroColumn-1] = "Nivå for første markering"
		$l_aGraphSheetArray[$g_iS3EventLevelRow-1][$g_iS3EventValueColumn-1] =$g_iS3Event_FirstLevelValue
		$l_aGraphSheetArray[$g_iS3EventDistanceRow-1][$g_iS3EventIntroColumn-1] ="Avstand mellom markeringer"
		$l_aGraphSheetArray[$g_iS3EventDistanceRow-1][$g_iS3EventValueColumn-1] =$g_iS3Event_DeltaLevelValue

		$l_aGraphSheetArray[$g_iS3ActivityHeaderRow-1][$g_iS3ActivityIntroColumn-1] = "Aktiviteter"
		$l_aGraphSheetArray[$g_iS3ActivityLevelRow -1][$g_iS3ActivityIntroColumn-1] = "Nivå for første markering"
		$l_aGraphSheetArray[$g_iS3ActivityLevelRow-1][$g_iS3ActivityValueColumn-1] =$g_iS3Activity_FirstLevelValue
		$l_aGraphSheetArray[$g_iS3ActivityDistanceRow-1][$g_iS3ActivityIntroColumn-1] ="Avstand mellom markeringer"
		$l_aGraphSheetArray[$g_iS3ActivityDistanceRow-1][$g_iS3ActivityValueColumn-1] =$g_iS3Activity_DeltaLevelValue

	Else
		$l_aGraphSheetArray[$g_iS3TimeSummaryHeaderRow-1][$g_iS3TimeSummaryHeaderColumn-1] = "Recorded period"
		$l_aGraphSheetArray[$g_iS3TimeSummaryHeaderRow-1][$g_iS3TimeSummaryFromColumn-1] = "From"
		$l_aGraphSheetArray[$g_iS3TimeSummaryHeaderRow-1][$g_iS3TimeSummaryToColumn-1] = "To"
		$l_aGraphSheetArray[$g_iS3GraphTimeResolutionRow-1][$g_iS3GraphTimeResolutionIntroColumn-1] = "Choose resolution for the main graph:"

		$l_aGraphSheetArray[$g_iS3GraphTimeResolutionRow-1][$g_iS3GraphTimeResolutionInputColumn-1] = "5"
		$l_aGraphSheetArray[$g_iS3GraphTimeResolutionRow-1][$g_iS3TimeSummaryToColumn-1] = " seconds"
		$l_aGraphSheetArray[$g_iS3GraphNumberOfDatapointsRow-1][$g_iS3GraphNumberOfDatapointsIntroColumn-1] = "Equals"
		$l_aGraphSheetArray[$g_iS3GraphNumberOfDatapointsRow-1][$g_iS3GraphNumberOfDatapointsValueColumn-1] = $l_sCreateTimeFormula
		$l_aGraphSheetArray[$g_iS3GraphNumberOfDatapointsRow-1][$g_iS3GraphNumberOfDatapointsText2Column-1] = "data points"

		$l_aGraphSheetArray[$g_iS3PulseSummaryHeaderRow-1][$g_iS3PulseSummaryIntroColumn-1] = "Summary: Heart Rate "
		$l_aGraphSheetArray[$g_iS3PulseSummaryAverageRow-1][$g_iS3PulseSummaryIntroColumn-1] = "Average"
		$l_aGraphSheetArray[$g_iS3PulseSummaryAverageRow-1][$g_iS3PulseSummaryValueColumn-1] = $l_sFormula_Average
		$l_aGraphSheetArray[$g_iS3PulseSummaryStDevRow-1][$g_iS3PulseSummaryIntroColumn-1] = "Standard deviation"
		$l_aGraphSheetArray[$g_iS3PulseSummaryStDevRow-1][$g_iS3PulseSummaryValueColumn-1] = $l_sFormula_StDev
		$l_aGraphSheetArray[$g_iS3PulseSummaryAvPlusOneStDevRow-1][$g_iS3PulseSummaryIntroColumn-1] = "Average plus one standard deviation"
		$l_aGraphSheetArray[$g_iS3PulseSummaryAvPlusOneStDevRow-1][$g_iS3PulseSummaryValueColumn-1] = $l_sFormula_AvPlusOneStDev
		$l_aGraphSheetArray[$g_iS3PulseSummaryAvPlusTwoStDevRow-1][$g_iS3PulseSummaryIntroColumn-1] ="Average plus two standard deviations"
		$l_aGraphSheetArray[$g_iS3PulseSummaryAvPlusTwoStDevRow-1][$g_iS3PulseSummaryValueColumn-1] = $l_sFormula_AvPlusTwoStDev

		$l_aGraphSheetArray[$g_iS3EventHeaderRow-1][$g_iS3EventIntroColumn-1] = "Events / Behavior"
		$l_aGraphSheetArray[$g_iS3EventLevelRow -1][$g_iS3EventIntroColumn-1] = "Level for first indicator"
		$l_aGraphSheetArray[$g_iS3EventLevelRow-1][$g_iS3EventValueColumn-1] =$g_iS3Event_FirstLevelValue
		$l_aGraphSheetArray[$g_iS3EventDistanceRow-1][$g_iS3EventIntroColumn-1] ="Distance between indicators"
		$l_aGraphSheetArray[$g_iS3EventDistanceRow-1][$g_iS3EventValueColumn-1] =$g_iS3Event_DeltaLevelValue

		$l_aGraphSheetArray[$g_iS3ActivityHeaderRow-1][$g_iS3ActivityIntroColumn-1] = "Activities"
		$l_aGraphSheetArray[$g_iS3ActivityLevelRow -1][$g_iS3ActivityIntroColumn-1] = "Level for first indicator"
		$l_aGraphSheetArray[$g_iS3ActivityLevelRow-1][$g_iS3ActivityValueColumn-1] =$g_iS3Activity_FirstLevelValue
		$l_aGraphSheetArray[$g_iS3ActivityDistanceRow-1][$g_iS3ActivityIntroColumn-1] ="Distance between indicators"
		$l_aGraphSheetArray[$g_iS3ActivityDistanceRow-1][$g_iS3ActivityValueColumn-1] =$g_iS3Activity_DeltaLevelValue
	EndIf


	If $g_bS3ShowResolutionWarning = True then
		If $g_sChosenLanguage = "Norsk" Then
			$l_aGraphSheetArray[$g_iS3GraphTimeResolutionRow-1][$g_iS3GraphTimeResolutionWarningColumn-1] = "NB Dersom du har for lav oppløsning kan det bli for mange datapunkter for Excel"
		Else
			$l_aGraphSheetArray[$g_iS3GraphTimeResolutionRow-1][$g_iS3GraphTimeResolutionWarningColumn-1] = ""
		EndIf
	Else
		$l_aGraphSheetArray[$g_iS3GraphTimeResolutionRow-1][$g_iS3GraphTimeResolutionWarningColumn-1] = " "
	EndIf

	$l_aGraphSheetArray[$g_iS3TimeSummaryContentRow-1][$g_iS3TimeSummaryFromColumn-1] = "='"& $g_sNameOfSheet_5_CalculationsMainGraph&"'!"&_Excel_ColumnToLetter($g_iS5Column_CorrectedTime)&$g_iS5FirstContentRow
	$l_aGraphSheetArray[$g_iS3TimeSummaryContentRow-1][$g_iS3TimeSummaryToColumn-1] = "='"& $g_sNameOfSheet_5_CalculationsMainGraph&"'!"&_Excel_ColumnToLetter($g_iS5Column_CorrectedTime)&$g_iS5LastContentRow

	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_3_GraphMain, $l_aGraphSheetArray, _Excel_ColumnToLetter($g_iS3TimeSummaryHeaderColumn) & $g_iS3TimeSummaryHeaderRow)
	$g_oExcel.Application.ScreenUpdating = True ;Turn Screen Updating on
EndFunc   ;==>_Fill_Sheet_3_MainGraph

Func _Add_MainGraph_To_Sheet_2()
	$g_oExcel.Application.ScreenUpdating = False ;Turn Screen Updating off

	Local $l_sXValueRange, $l_asDataRange, $l_asDataName, $l_sChartAreaRange, $l_sChartHeading

	;Legger til hovedgraf på side $g_sNameOfSheet_3_GraphMain
	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), " ", "Lager hovedgraf")
		$l_sChartHeading = 'Puls (Hjerteslag per minutt)'
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), " ", "Making main graph")
		$l_sChartHeading = 'Heart Rate (beats per minute)'
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1



	;Velger datapunkter fra $g_sNameOfSheet_5_CalculationsMainGraph
	$l_sXValueRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Column_TimeForLongGraph) & $g_iS5FirstContentRow & ":" & _Excel_ColumnToLetter($g_iS5Column_TimeForLongGraph) & $g_iS5LastSecondIntervalRow
	$l_asDataRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Column_PulseForLongGraph) & $g_iS5FirstContentRow & ":" & _Excel_ColumnToLetter($g_iS5Column_PulseForLongGraph) & $g_iS5LastSecondIntervalRow
	$l_asDataName = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Column_Pulse) & $g_iS5MainHeaderRow

	Local $l_iNumberOfColumnsToSpan = Round($g_iNumberOfSeconds / 50, 0)
	If $l_iNumberOfColumnsToSpan < 15 Then
		$l_iNumberOfColumnsToSpan = 15
	EndIf
	Local $l_iNumberOfColumnsToSpan = 25
	$l_sChartAreaRange = _Excel_ColumnToLetter($g_iS3GraphFirstColumn) & $g_iS3GraphFirstRow & ":" & _Excel_ColumnToLetter($g_iS3GraphFirstColumn+$l_iNumberOfColumnsToSpan) & $g_iS3GraphFirstRow+$l_iNumberOfColumnsToSpan
	;_XLChart_ChartCreate($oExcel, $vWorksheet, $iChartType, $sSizeByCells, $sChartName, $sXValueRange, $vDataRange, $vDataName[, $bShowLegend = True[, $sTitle = ""[, $sXTitle = ""[, $sYTitle = ""[, $sZTitle = ""[, $bShowDatatable = False[, $bScreenUpdate = False]]]]]]])
	Local $l_oChart = _XLChart_ChartCreate($g_oExcel, $g_sNameOfSheet_3_GraphMain, $xlXYScatterLines,  $l_sChartAreaRange, $l_sChartHeading, $l_sXValueRange, $l_asDataRange, $l_asDataName, True, $l_sChartHeading)
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "Excel: _ChartCreate", "Error creating chart")
	EndIf

	;Find time of first datapoint
	Local $l_iFirstTime = $g_iHourMin + $g_iMinuteMin / 60
	;$l_iFirstTime = $l_iFirstTime - 1 / 60 ; Set first point 1 minute before first trackpoint

	Local $l_iLastTime = $g_iHourMax + $g_iMinuteMax / 60
	$l_iLastTime = $l_iLastTime + 1 / 60 ; Set last point 1 minute after last trackpoint

	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Justerer x-akse", "Lager hovedgraf ")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Adjusting x-axis", "Making main graph ")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1

	_XLChart_AxisSet($l_oChart.Axes(1), $l_iFirstTime / 24, $l_iLastTime / 24)
	With $l_oChart.Legend
		.Position = $xlLeft
		.IncludeInLayout = True
	EndWith
	With $l_oChart.Axes(1)
		If ($l_iLastTime - $l_iFirstTime) < 1 Then
			.MajorUnit = 1 / (24 * 6)
			.MinorUnit = 1 / (24 * 60 * 60)
		Else
			.MajorUnit = 1 / (24)
			.MinorUnit = 1 / (24 * 6)
		EndIf
	EndWith
	$g_oExcel.Application.ScreenUpdating = True ;Turn Screen Updating back on

	;Legge til Puls-gjennomsnitt, -standardavvik og -2*standardavvik
	Local $l_iStatisticsSeries
	;Gjennomsnitt
	$l_asDataName = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header) & $g_iS5Row_LongGraphAverage_Header
	$l_sXValueRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header)& $g_iS5Row_LongGraphAverage_FirstValue &":"& _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header)& $g_iS5Row_LongGraphAverage_LastValue
	$l_asDataRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Value)& $g_iS5Row_LongGraphAverage_FirstValue &":"& _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Value)& $g_iS5Row_LongGraphAverage_LastValue

	$l_iStatisticsSeries = _XLChart_SeriesAdd($l_oChart, $l_sXValueRange, $l_asDataRange, $l_asDataName)
	With $l_iStatisticsSeries
		.Smooth = False
		.Format.Line.Visible = True
		.Format.Line.ForeColor.RGB = $COLOR_BLACK
		.MarkerStyle = -4142 ; xlMarkerStyleNone =	-4142
		.HasDataLabels = False
	EndWith

	;Gjennomsnitt pluss standardavvik
	$l_asDataName = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header) & $g_iS5Row_LongGraphOneStDev_Header
	$l_sXValueRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header)& $g_iS5Row_LongGraphOneStDev_FirstValue &":"& _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header)& $g_iS5Row_LongGraphOneStDev_LastValue
	$l_asDataRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Value)& $g_iS5Row_LongGraphOneStDev_FirstValue &":"& _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Value)& $g_iS5Row_LongGraphOneStDev_LastValue

	$l_iStatisticsSeries = _XLChart_SeriesAdd($l_oChart, $l_sXValueRange, $l_asDataRange, $l_asDataName)
	With $l_iStatisticsSeries
		.Smooth = False
		.Format.Line.Visible = True
		.Format.Line.ForeColor.RGB = $COLOR_FUCHSIA
		.MarkerStyle = -4142 ; xlMarkerStyleNone =	-4142
		.HasDataLabels = False
	EndWith

	;Gjennomsnitt pluss to standardavvik
	$l_asDataName = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header) & $g_iS5Row_LongGraphTwoStDev_Header
	$l_sXValueRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header)& $g_iS5Row_LongGraphTwoStDev_FirstValue &":"& _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header)& $g_iS5Row_LongGraphTwoStDev_LastValue
	$l_asDataRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Value)& $g_iS5Row_LongGraphTwoStDev_FirstValue &":"& _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Value)& $g_iS5Row_LongGraphTwoStDev_LastValue

	$l_iStatisticsSeries = _XLChart_SeriesAdd($l_oChart, $l_sXValueRange, $l_asDataRange, $l_asDataName)
	With $l_iStatisticsSeries
		.Smooth = False
		.Format.Line.Visible = True
		.Format.Line.ForeColor.RGB = $COLOR_RED
		.MarkerStyle = -4142 ; xlMarkerStyleNone =	-4142
		.HasDataLabels = False
	EndWith


	;_XLChart_LegendSet($l_oChart, Default, $xlLeft,Default, Default, Default, Default, Default)
	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Legger inn Aktivitetsmarkeringer", "Lager hovedgraf ")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Adding Activity marks", "Making main graph ")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	$g_oExcel.Application.ScreenUpdating = False ;Turn Screen Updating off

	;Legger til Aktivitetsmerker
	Local $l_oActivity_series, $l_iNumberOfVisibleActivities
	$l_iNumberOfVisibleActivities = 5


	;Local $l_sXActivityValueRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5FirstContentRow & ":" & _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5LastSecondIntervalRow
	Local $l_sXActivityValueRange
	Local $l_vActivityDataRange
	Local $l_vActivityNameRange

	$l_iCntr = 0
	While $l_iCntr < $g_iNumberOfActivities
		If $g_sChosenLanguage = "Norsk" Then
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Legger inn Aktivitetsmarkeringer: " & $l_iCntr + 1 & "/" & $g_iNumberOfActivities, "Lager hovedgraf ")
		Else
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Adding Activity marks: " & $l_iCntr + 1 & "/" & $g_iNumberOfActivities, "Making main graph ")
		EndIf
		$g_iProgressCounter = $g_iProgressCounter + 1


		$l_vActivityNameRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Activity_ForLong_ValueColumn) & $g_iS5Activity_ForLong_FirstHeaderRow + ($l_iCntr*$g_iS5Activity_ForLong_NumberOfRowsInSubTable)
		$l_sXActivityValueRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Activity_ForLong_TimeColumn)& $g_iS5Activity_ForLong_FirstContentRow +($l_iCntr*$g_iS5Activity_ForLong_NumberOfRowsInSubTable) &":"& _Excel_ColumnToLetter($g_iS5Activity_ForLong_TimeColumn)& $g_iS5Activity_ForLong_FirstContentRow +($l_iCntr*$g_iS5Activity_ForLong_NumberOfRowsInSubTable) +1
		$l_vActivityDataRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Activity_ForLong_ValueColumn)& $g_iS5Activity_ForLong_FirstContentRow +($l_iCntr*$g_iS5Activity_ForLong_NumberOfRowsInSubTable) &":"& _Excel_ColumnToLetter($g_iS5Activity_ForLong_ValueColumn)& $g_iS5Activity_ForLong_FirstContentRow +($l_iCntr*$g_iS5Activity_ForLong_NumberOfRowsInSubTable) +1


		$l_oActivity_series = _XLChart_SeriesAdd($l_oChart, $l_sXActivityValueRange, $l_vActivityDataRange, $l_vActivityNameRange)

		With $l_oActivity_series
			.Smooth = False
			.Format.Line.Visible = True
			.Format.Line.Weight = 3.5
			.MarkerStyle = 1 ; 1= xlMarkerStyleSquare
			.HasDataLabels = False
			If $l_iCntr >= $l_iNumberOfVisibleActivities Then
				.IsFiltered = True
			EndIf
		EndWith
		$l_iCntr = $l_iCntr + 1
	WEnd
	$g_oExcel.Application.ScreenUpdating = True ;Turn Screen Updating back on

	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Legger inn Hendelsesmarkeringer", "Lager hovedgraf ")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Adding Event marks", "Making main graph ")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	;Legger til Hendelsessmerker
	Local $l_oEvent_series, $l_iNumberOfVisibleEvents
	$l_iNumberOfVisibleEvents = 5
	$g_oExcel.Application.ScreenUpdating = False ;Turn Screen Updating off
	$l_iCntr = 0

	;Local $l_sXEventValueRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5FirstContentRow & ":" & _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5LastSecondIntervalRow
	Local $l_sXEventValueRange
	Local $l_vEventDataRange
	Local $l_vEventNameRange

	While $l_iCntr < $g_iNumberOfEvents
		If $g_sChosenLanguage = "Norsk" Then
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Legger inn Hendelsesmarkeringer: " & $l_iCntr + 1 & "/" & $g_iNumberOfEvents, "Lager hovedgraf ")
		Else
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Adding Event marks: " & $l_iCntr + 1 & "/" & $g_iNumberOfEvents, "Making main graph ")
		EndIf
		$g_iProgressCounter = $g_iProgressCounter + 1

		$l_vEventNameRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Event_ForLong_ValueColumn) & $g_iS5Event_ForLong_FirstHeaderRow + ($l_iCntr*$g_iS5Event_ForLong_NumberOfRowsInSubTable)
		$l_sXEventValueRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Event_ForLong_TimeColumn)& $g_iS5Event_ForLong_FirstContentRow +($l_iCntr*$g_iS5Event_ForLong_NumberOfRowsInSubTable) &":"& _Excel_ColumnToLetter($g_iS5Event_ForLong_TimeColumn)& $g_iS5Event_ForLong_FirstContentRow +($l_iCntr*$g_iS5Event_ForLong_NumberOfRowsInSubTable) +1
		$l_vEventDataRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Event_ForLong_ValueColumn)& $g_iS5Event_ForLong_FirstContentRow +($l_iCntr*$g_iS5Event_ForLong_NumberOfRowsInSubTable) &":"& _Excel_ColumnToLetter($g_iS5Event_ForLong_ValueColumn)& $g_iS5Event_ForLong_FirstContentRow +($l_iCntr*$g_iS5Event_ForLong_NumberOfRowsInSubTable) +1

		$l_oEvent_series = _XLChart_SeriesAdd($l_oChart, $l_sXEventValueRange, $l_vEventDataRange, $l_vEventNameRange)
		; "='"&$g_sNameOfSheet_5_CalculationsMainGraph&"'!R5C"&$l_iStartColumnEvent+$l_iCntr&":R"&$l_iLastRow&"C"&$l_iStartColumnEvent+$l_iCntr, "='"&$g_sNameOfSheet_5_CalculationsMainGraph&"'!R4C"&$l_iStartColumnEvent+$l_iCntr)
		With $l_oEvent_series
			.Smooth = False
			.Format.Line.Visible = False
			.MarkerStyle = 8 ; xlMarkerStyleCircle	8	Circular markers
			.MarkerSize = 10 ; min = 2, max = 72
			.HasDataLabels = True
			.DataLabels.ShowValue = False
			.DataLabels.ShowSeriesName = True
			.DataLabels.Position = 0 ;xlLabelPositionAbove	 = 0	Data label is positioned above the data point.
			.DataLabels.Orientation = -4171 ; xlUpward= 	-4171	Text runs upward.
			If $l_iCntr >= $l_iNumberOfVisibleEvents Then
				.IsFiltered = True
			EndIf
		EndWith

		$l_iCntr = $l_iCntr + 1
	WEnd

	$g_oExcel.Application.ScreenUpdating = True ;Turn Screen Updating on
EndFunc   ;==>_Add_MainGraph_To_Sheet_2


Func _Fill_Sheet_4_GraphSelection()
	$g_oExcel.Application.ScreenUpdating = False ;Turn Screen Updating off
	;$g_sNameOfSheet_2_Observations: Inserting time range header

	Local $l_sCreateTimeFormula = "=ROUNDDOWN(COUNT('" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!"& _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5FirstContentRow & ":" & _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5LastSecondIntervalRow&")/$"& _Excel_ColumnToLetter($g_iS4GraphNumberOfDatapointsValueColumn) & "$" & $g_iS4GraphTimeResolutionRow &";0)"

	Local $l_sShortPulseRange = "'"&$g_sNameOfSheet_5_CalculationsMainGraph&"'!$"& _Excel_ColumnToLetter($g_iS5Column_PulseForShortGraph) & "$"& $g_iS5FirstContentRow &":$" & _Excel_ColumnToLetter($g_iS5Column_PulseForShortGraph) & "$"& $g_iS5LastSecondIntervalRow
	Local $l_sFormula_Average =  "=AVERAGEIF("&$l_sShortPulseRange  & ";"& Chr(34) &"<>#N/A"&Chr(34)&")"
	Local $l_sFormula_StDev = "=AGGREGATE( 8;6;"&$l_sShortPulseRange  & ")"
	Local $l_sFormula_AvPlusOneStDev = "=AGGREGATE( 8;6;"&$l_sShortPulseRange  & ") + " &_Excel_ColumnToLetter($g_iS4PulseSummaryValueColumn) & $g_iS4PulseSummaryAverageRow
	Local $l_sFormula_AvPlusTwoStDev = "=2*AGGREGATE( 8;6;"&$l_sShortPulseRange  & ") + " &_Excel_ColumnToLetter($g_iS4PulseSummaryValueColumn) & $g_iS4PulseSummaryAverageRow


	Local $l_iCntr = 0
	Local $l_iGraphSheetTableWidth = $g_iS4AxisStopColumn
	Local $l_iGraphSheetTableHeight = $g_iS4ActivityDistanceRow + 3
	Local $l_aGraphSheetArray[$l_iGraphSheetTableHeight][$l_iGraphSheetTableWidth]

	IF $g_sChosenLanguage = "Norsk" Then
		$l_aGraphSheetArray[$g_iS4TimeSummaryHeaderRow-1][$g_iS4TimeSummaryHeaderColumn-1] = "Måleperiode"
		$l_aGraphSheetArray[$g_iS4TimeSummaryHeaderRow-1][$g_iS4TimeSummaryFromColumn-1] = "Fra kl"
		$l_aGraphSheetArray[$g_iS4TimeSummaryHeaderRow-1][$g_iS4TimeSummaryToColumn-1] = "Til kl"
		$l_aGraphSheetArray[$g_iS4GraphTimeSelectionRow-1][$g_iS4GraphTimeSelectionIntroColumn-1] = "Velg periode for grafutsnitt"
		$l_aGraphSheetArray[$g_iS4GraphTimeResolutionRow-1][$g_iS4GraphTimeResolutionIntroColumn-1] = "Velg oppløsning for hovedgrafen:"
		If $g_bS4ShowResolutionWarning = True Then
			$l_aGraphSheetArray[$g_iS4GraphTimeResolutionRow-1][$g_iS4GraphTimeResolutionWarningColumn-1] = "NB Dersom du har for lav oppløsning kan det bli for mange datapunkter for Excel"
		Else
			$l_aGraphSheetArray[$g_iS4GraphTimeResolutionRow-1][$g_iS4GraphTimeResolutionWarningColumn-1] = " "
		EndIf
		$l_aGraphSheetArray[$g_iS4GraphTimeResolutionRow-1][$g_iS4GraphTimeResolutionInputColumn-1] = "1"
		$l_aGraphSheetArray[$g_iS4GraphTimeResolutionRow-1][$g_iS4TimeSummaryToColumn-1] = " sekunder"
		$l_aGraphSheetArray[$g_iS4GraphNumberOfDatapointsRow-1][$g_iS4GraphNumberOfDatapointsIntroColumn-1] = "Tilsvarer"
		$l_aGraphSheetArray[$g_iS4GraphNumberOfDatapointsRow-1][$g_iS4GraphNumberOfDatapointsValueColumn-1] = $l_sCreateTimeFormula
		$l_aGraphSheetArray[$g_iS4GraphNumberOfDatapointsRow-1][$g_iS4GraphNumberOfDatapointsText2Column-1] = "datapunkter"

		$l_aGraphSheetArray[$g_iS4TimeSummaryContentRow-1][$g_iS4TimeSummaryFromColumn-1] = "='"& $g_sNameOfSheet_5_CalculationsMainGraph&"'!"&_Excel_ColumnToLetter($g_iS5Column_CorrectedTime)&$g_iS5FirstContentRow
		$l_aGraphSheetArray[$g_iS4TimeSummaryContentRow-1][$g_iS4TimeSummaryToColumn-1] = "='"& $g_sNameOfSheet_5_CalculationsMainGraph&"'!"&_Excel_ColumnToLetter($g_iS5Column_CorrectedTime)&$g_iS5LastContentRow

		$l_aGraphSheetArray[$g_iS4PulseSummaryHeaderRow-1][$g_iS4PulseSummaryIntroColumn-1] = "Oppsummering puls "
		$l_aGraphSheetArray[$g_iS4PulseSummaryHeaderExtraInfoRow -1][$g_iS4PulseSummaryIntroColumn-1] = "(Gjelder valgt utsnitt)"
		$l_aGraphSheetArray[$g_iS4PulseSummaryAverageRow-1][$g_iS4PulseSummaryIntroColumn-1] = "Gjennomsnitt"
		$l_aGraphSheetArray[$g_iS4PulseSummaryAverageRow-1][$g_iS4PulseSummaryValueColumn-1] = $l_sFormula_Average
		$l_aGraphSheetArray[$g_iS4PulseSummaryStDevRow-1][$g_iS4PulseSummaryIntroColumn-1] = "Standardavvik"
		$l_aGraphSheetArray[$g_iS4PulseSummaryStDevRow-1][$g_iS4PulseSummaryValueColumn-1] = $l_sFormula_StDev
		$l_aGraphSheetArray[$g_iS4PulseSummaryAvPlusOneStDevRow-1][$g_iS4PulseSummaryIntroColumn-1] = "Gjennomsnitt pluss ett standardavvik"
		$l_aGraphSheetArray[$g_iS4PulseSummaryAvPlusOneStDevRow-1][$g_iS4PulseSummaryValueColumn-1] = $l_sFormula_AvPlusOneStDev
		$l_aGraphSheetArray[$g_iS4PulseSummaryAvPlusTwoStDevRow-1][$g_iS4PulseSummaryIntroColumn-1] ="Gjennomsnitt pluss to standardavvik"
		$l_aGraphSheetArray[$g_iS4PulseSummaryAvPlusTwoStDevRow-1][$g_iS4PulseSummaryValueColumn-1] = $l_sFormula_AvPlusTwoStDev

		$l_aGraphSheetArray[$g_iS4EventHeaderRow-1][$g_iS4EventIntroColumn-1] = "Hendelser"
		$l_aGraphSheetArray[$g_iS4EventLevelRow -1][$g_iS4EventIntroColumn-1] = "Nivå for første markering"
		$l_aGraphSheetArray[$g_iS4EventLevelRow-1][$g_iS4EventValueColumn-1] =$g_iS4Event_FirstLevelValue
		$l_aGraphSheetArray[$g_iS4EventDistanceRow-1][$g_iS4EventIntroColumn-1] ="Avstand mellom markeringer"
		$l_aGraphSheetArray[$g_iS4EventDistanceRow-1][$g_iS4EventValueColumn-1] =$g_iS4Event_DeltaLevelValue

		$l_aGraphSheetArray[$g_iS4ActivityHeaderRow-1][$g_iS4ActivityIntroColumn-1] = "Aktiviteter"
		$l_aGraphSheetArray[$g_iS4ActivityLevelRow -1][$g_iS4ActivityIntroColumn-1] = "Nivå for første markering"
		$l_aGraphSheetArray[$g_iS4ActivityLevelRow-1][$g_iS4ActivityValueColumn-1] =$g_iS4Activity_FirstLevelValue
		$l_aGraphSheetArray[$g_iS4ActivityDistanceRow-1][$g_iS4ActivityIntroColumn-1] ="Avstand mellom markeringer"
		$l_aGraphSheetArray[$g_iS4ActivityDistanceRow-1][$g_iS4ActivityValueColumn-1] =$g_iS4Activity_DeltaLevelValue


		$l_aGraphSheetArray[$g_sS4AxisWarningRow-1][$g_iS4AxisWarningColumn-1] = $g_sS4AxisWarning
		$l_aGraphSheetArray[$g_sS4AxisInfoRow-1][$g_iS4AxisStartColumn-1] = "Akse start"
		$l_aGraphSheetArray[$g_sS4AxisInfoRow-1][$g_iS4AxisStopColumn-1] = "Akse stopp"
	Else
		$l_aGraphSheetArray[$g_iS4TimeSummaryHeaderRow-1][$g_iS4TimeSummaryHeaderColumn-1] = "Recorded period"
		$l_aGraphSheetArray[$g_iS4TimeSummaryHeaderRow-1][$g_iS4TimeSummaryFromColumn-1] = "From"
		$l_aGraphSheetArray[$g_iS4TimeSummaryHeaderRow-1][$g_iS4TimeSummaryToColumn-1] = "To"
		$l_aGraphSheetArray[$g_iS4GraphTimeSelectionRow-1][$g_iS4GraphTimeSelectionIntroColumn-1] = "Selected time period for graph range"
		$l_aGraphSheetArray[$g_iS4GraphTimeResolutionRow-1][$g_iS4GraphTimeResolutionIntroColumn-1] = "Choose resolution for the selected graph range:"

		$l_aGraphSheetArray[$g_iS4GraphTimeResolutionRow-1][$g_iS4GraphTimeResolutionWarningColumn-1] = " "

		$l_aGraphSheetArray[$g_iS4GraphTimeResolutionRow-1][$g_iS4GraphTimeResolutionInputColumn-1] = "1"
		$l_aGraphSheetArray[$g_iS4GraphTimeResolutionRow-1][$g_iS4TimeSummaryToColumn-1] = " seconds"
		$l_aGraphSheetArray[$g_iS4GraphNumberOfDatapointsRow-1][$g_iS4GraphNumberOfDatapointsIntroColumn-1] = "Equals"
		$l_aGraphSheetArray[$g_iS4GraphNumberOfDatapointsRow-1][$g_iS4GraphNumberOfDatapointsValueColumn-1] = $l_sCreateTimeFormula
		$l_aGraphSheetArray[$g_iS4GraphNumberOfDatapointsRow-1][$g_iS4GraphNumberOfDatapointsText2Column-1] = "data points"

		$l_aGraphSheetArray[$g_iS4TimeSummaryContentRow-1][$g_iS4TimeSummaryFromColumn-1] = "='"& $g_sNameOfSheet_5_CalculationsMainGraph&"'!"&_Excel_ColumnToLetter($g_iS5Column_CorrectedTime)&$g_iS5FirstContentRow
		$l_aGraphSheetArray[$g_iS4TimeSummaryContentRow-1][$g_iS4TimeSummaryToColumn-1] = "='"& $g_sNameOfSheet_5_CalculationsMainGraph&"'!"&_Excel_ColumnToLetter($g_iS5Column_CorrectedTime)&$g_iS5LastContentRow

		$l_aGraphSheetArray[$g_iS4PulseSummaryHeaderRow-1][$g_iS4PulseSummaryIntroColumn-1] = "Summary: Heart Rate"
		$l_aGraphSheetArray[$g_iS4PulseSummaryHeaderExtraInfoRow -1][$g_iS4PulseSummaryIntroColumn-1] = "(For selected time range)"
		$l_aGraphSheetArray[$g_iS4PulseSummaryAverageRow-1][$g_iS4PulseSummaryIntroColumn-1] = "Average"
		$l_aGraphSheetArray[$g_iS4PulseSummaryAverageRow-1][$g_iS4PulseSummaryValueColumn-1] = $l_sFormula_Average
		$l_aGraphSheetArray[$g_iS4PulseSummaryStDevRow-1][$g_iS4PulseSummaryIntroColumn-1] = "Standard deviation"
		$l_aGraphSheetArray[$g_iS4PulseSummaryStDevRow-1][$g_iS4PulseSummaryValueColumn-1] = $l_sFormula_StDev
		$l_aGraphSheetArray[$g_iS4PulseSummaryAvPlusOneStDevRow-1][$g_iS4PulseSummaryIntroColumn-1] = "Average plus one standard deviation"
		$l_aGraphSheetArray[$g_iS4PulseSummaryAvPlusOneStDevRow-1][$g_iS4PulseSummaryValueColumn-1] = $l_sFormula_AvPlusOneStDev
		$l_aGraphSheetArray[$g_iS4PulseSummaryAvPlusTwoStDevRow-1][$g_iS4PulseSummaryIntroColumn-1] ="Average plus two standard deviations"
		$l_aGraphSheetArray[$g_iS4PulseSummaryAvPlusTwoStDevRow-1][$g_iS4PulseSummaryValueColumn-1] = $l_sFormula_AvPlusTwoStDev

		$l_aGraphSheetArray[$g_iS4EventHeaderRow-1][$g_iS4EventIntroColumn-1] = "Events / Behavior"
		$l_aGraphSheetArray[$g_iS4EventLevelRow -1][$g_iS4EventIntroColumn-1] = "Level for first indicator"
		$l_aGraphSheetArray[$g_iS4EventLevelRow-1][$g_iS4EventValueColumn-1] =$g_iS4Event_FirstLevelValue
		$l_aGraphSheetArray[$g_iS4EventDistanceRow-1][$g_iS4EventIntroColumn-1] ="Distance between indicators"
		$l_aGraphSheetArray[$g_iS4EventDistanceRow-1][$g_iS4EventValueColumn-1] =$g_iS4Event_DeltaLevelValue

		$l_aGraphSheetArray[$g_iS4ActivityHeaderRow-1][$g_iS4ActivityIntroColumn-1] = "Activities"
		$l_aGraphSheetArray[$g_iS4ActivityLevelRow -1][$g_iS4ActivityIntroColumn-1] = "Level for first indicator"
		$l_aGraphSheetArray[$g_iS4ActivityLevelRow-1][$g_iS4ActivityValueColumn-1] =$g_iS4Activity_FirstLevelValue
		$l_aGraphSheetArray[$g_iS4ActivityDistanceRow-1][$g_iS4ActivityIntroColumn-1] ="Distance between indicators"
		$l_aGraphSheetArray[$g_iS4ActivityDistanceRow-1][$g_iS4ActivityValueColumn-1] =$g_iS4Activity_DeltaLevelValue


		$l_aGraphSheetArray[$g_sS4AxisWarningRow-1][$g_iS4AxisWarningColumn-1] = $g_sS4AxisWarning
		$l_aGraphSheetArray[$g_sS4AxisInfoRow-1][$g_iS4AxisStartColumn-1] = "Axis start"
		$l_aGraphSheetArray[$g_sS4AxisInfoRow-1][$g_iS4AxisStopColumn-1] = "Axis stop"
	EndIf

	;=(HOUR(D3)+MINUTE(D3)/60)/24
	Local $l_sAxisFormula = "=(HOUR(" & _Excel_ColumnToLetter($g_iS4GraphTimeSelectionFromColumn) & $g_iS4GraphTimeSelectionRow &")+MINUTE("& _
		_Excel_ColumnToLetter($g_iS4GraphTimeSelectionFromColumn) & $g_iS4GraphTimeSelectionRow &")/60)/24"
	$l_aGraphSheetArray[$g_sS4AxisValueRow-1][$g_iS4AxisStartColumn-1] = $l_sAxisFormula
	$l_sAxisFormula = "=(HOUR(" & _Excel_ColumnToLetter($g_iS4GraphTimeSelectionToColumn) & $g_iS4GraphTimeSelectionRow &")+MINUTE("& _
		_Excel_ColumnToLetter($g_iS4GraphTimeSelectionToColumn) & $g_iS4GraphTimeSelectionRow &")/60)/24"
	$l_aGraphSheetArray[$g_sS4AxisValueRow-1][$g_iS4AxisStopColumn-1] = $l_sAxisFormula

	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_4_GraphSelection, $l_aGraphSheetArray, _Excel_ColumnToLetter($g_iS4TimeSummaryHeaderColumn) & $g_iS4TimeSummaryHeaderRow)
	$g_oExcel.Application.ScreenUpdating = True ;Turn Screen Updating on

EndFunc   ;==>_Fill_Sheet_4_GraphSelection()


Func _Format_Sheet_4_GraphSelection()
	$g_oExcel.Application.ScreenUpdating = False ;Turn Screen Updating off
	;	Time symmary
	;Header
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4TimeSummaryHeaderColumn) & $g_iS4TimeSummaryHeaderRow & ":" & _Excel_ColumnToLetter($g_iS4TimeSummaryToColumn) & $g_iS4TimeSummaryHeaderRow)
		.Font.Bold = True
		.Font.Color = $g_iTableMainHeaderFontColor
		.Font.Size = $g_iTableMainHeaderFontSize
		.Interior.Color = $g_iTableMainHeaderColorMaster
	EndWith
	;All introtext in first columns
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4TimeSummaryHeaderColumn) & $g_iS4TimeSummaryContentRow & ":" & _Excel_ColumnToLetter($g_iS4TimeSummaryFromColumn - 1) & $g_iS4GraphNumberOfDatapointsRow)
		;.MergeCells = True
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableMainHeaderColorMaster
	EndWith
	;Introtext for data resolution
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4GraphTimeResolutionIntroColumn) & $g_iS4GraphTimeResolutionRow & ":" & _Excel_ColumnToLetter($g_iS4TimeSummaryFromColumn - 1) & $g_iS4GraphTimeResolutionRow)
		.Merge;Cells = True
		.WrapText = True
		.ShrinkToFit = False
	EndWith
	;Time summary contents
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4TimeSummaryFromColumn) & $g_iS4TimeSummaryContentRow & ":" & _Excel_ColumnToLetter($g_iS4TimeSummaryToColumn) & $g_iS4TimeSummaryContentRow)
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableSecondHeaderColorMaster
	EndWith

	;Unused To-column-cells
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4TimeSummaryToColumn) & $g_iS4GraphTimeResolutionRow & ":" & _Excel_ColumnToLetter($g_iS4TimeSummaryToColumn) & $g_iS4GraphNumberOfDatapointsRow )
		;.MergeCells = True
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableMainHeaderColorMaster
	EndWith

	;Number of datapoints
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4GraphNumberOfDatapointsValueColumn) & $g_iS4GraphNumberOfDatapointsRow)
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableSecondHeaderColorMaster
	EndWith

	;Borders around
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4TimeSummaryHeaderColumn) & $g_iS4TimeSummaryHeaderRow & ":" & _Excel_ColumnToLetter($g_iS4TimeSummaryToColumn) & $g_iS4GraphTimeResolutionRow)
		;.WrapText = True
		.Borders.LineStyle = $g_xlContinuous
		.Borders.Weight = $xlThick
		.Borders.Color = $g_iTableMainHeaderColorMaster
	EndWith

	If $g_bS4ShowResolutionWarning = True Then
		;Warning in yellow
		With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4GraphTimeResolutionWarningColumn) & $g_iS4GraphTimeResolutionRow & ":" & _Excel_ColumnToLetter($g_iS4GraphTimeResolutionWarningColumn + 5) & $g_iS4GraphTimeResolutionRow)
			.Merge ;Cells = True
			.WrapText = True
			.Font.Bold = True
			.Interior.Color = $g_iPrivateColor_MEDIUM_YELLOW
		EndWith
	EndIf

	With $g_oWorkbook.Worksheets($g_sNameOfSheet_4_GraphSelection).Rows($g_iS4GraphTimeResolutionRow)
        .RowHeight = 30
    EndWith

	;Format Axis summary
	;Warning in yellow
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4AxisWarningColumn) & $g_sS4AxisWarningRow & ":" & _Excel_ColumnToLetter($g_iS4AxisWarningColumn + 10) & $g_sS4AxisWarningRow)
		.Merge ;Cells = True
		;.WrapText = True
		;.Font.Size = $g_iTableMainHeaderFontSize
		.Font.Bold = True
		.Interior.Color = $g_iPrivateColor_MEDIUM_YELLOW
	EndWith
	;All introtext for axes
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4AxisStartColumn) & $g_sS4AxisInfoRow & ":" & _Excel_ColumnToLetter($g_iS4AxisStopColumn) & $g_sS4AxisInfoRow)
		;.MergeCells = True
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableMainHeaderColorMaster
	EndWith
	;Start and stop values
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4AxisStartColumn) & $g_sS4AxisValueRow & ":" & _Excel_ColumnToLetter($g_iS4AxisStopColumn) & $g_sS4AxisValueRow)
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableSecondHeaderColorMaster
	EndWith
	;Borders around
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4AxisStartColumn) & $g_sS4AxisInfoRow & ":" & _Excel_ColumnToLetter($g_iS4AxisStopColumn) & $g_sS4AxisValueRow)
		;.WrapText = True
		.Borders.LineStyle = $g_xlContinuous
		.Borders.Weight = $xlThick
		.Borders.Color = $g_iTableMainHeaderColorMaster
	EndWith

	;Format Pulse summary
	;Header
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4PulseSummaryIntroColumn) & $g_iS4PulseSummaryHeaderRow & ":" & _Excel_ColumnToLetter($g_iS4TimeSummaryToColumn) & $g_iS4PulseSummaryHeaderExtraInfoRow)
		.Font.Bold = True
		.Font.Color = $g_iTableMainHeaderFontColor
		.Font.Size = $g_iTableMainHeaderFontSize
		.Interior.Color = $g_iTableMainHeaderColorMaster
	EndWith
	;All introtext in first columns
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4PulseSummaryIntroColumn) & $g_iS4PulseSummaryAverageRow & ":" & _Excel_ColumnToLetter($g_iS4TimeSummaryFromColumn) & $g_iS4PulseSummaryAvPlusTwoStDevRow)
		;.MergeCells = True
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableMainHeaderColorMaster
	EndWith

	;Pulse summary Values
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4PulseSummaryValueColumn) & $g_iS4PulseSummaryAverageRow & ":" & _Excel_ColumnToLetter($g_iS4PulseSummaryValueColumn) & $g_iS4PulseSummaryAvPlusTwoStDevRow)
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableSecondHeaderColorMaster
		.NumberFormat = "0,00"
	EndWith

	;Borders around
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4PulseSummaryIntroColumn) & $g_iS4PulseSummaryHeaderRow & ":" & _Excel_ColumnToLetter($g_iS4PulseSummaryValueColumn) & $g_iS4PulseSummaryAvPlusTwoStDevRow)
		;.WrapText = True
		.Borders.LineStyle = $g_xlContinuous
		.Borders.Weight = $xlThick
		.Borders.Color = $g_iTableMainHeaderColorMaster
	EndWith


	;Format Activity
	;	Green header
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4ActivityIntroColumn) & $g_iS4ActivityHeaderRow & ":" & _Excel_ColumnToLetter($g_iS4ActivityValueColumn) & $g_iS4ActivityHeaderRow)
		.MergeCells = True
		.Font.Bold = True
		.Font.Color = $g_iTableMainHeaderFontColor
		.Font.Size = $g_iTableMainHeaderFontSize
		.Interior.Color = $g_iTableMainHeaderColorActivity
	EndWith
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4ActivityIntroColumn) & $g_iS4ActivityLevelRow & ":" & _Excel_ColumnToLetter($g_iS4ActivityValueColumn - 1) & $g_iS4ActivityLevelRow)
		.MergeCells = True
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableMainHeaderColorActivity
	EndWith
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4ActivityIntroColumn) & $g_iS4ActivityDistanceRow & ":" & _Excel_ColumnToLetter($g_iS4ActivityValueColumn - 1) & $g_iS4ActivityDistanceRow)
		.MergeCells = True
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableMainHeaderColorActivity
	EndWith
	;Borders around input cells
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4ActivityValueColumn) & $g_iS4ActivityLevelRow & ":" & _Excel_ColumnToLetter($g_iS4ActivityValueColumn) & $g_iS4ActivityDistanceRow )
		.WrapText = True
		.Borders.LineStyle = $g_xlContinuous
		.Borders.Weight = $xlThick
		.Borders.Color = $g_iTableMainHeaderColorActivity
	EndWith

	;Format Event
	;	Yellow header
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4EventIntroColumn) & $g_iS4EventHeaderRow & ":" & _Excel_ColumnToLetter($g_iS4EventValueColumn) & $g_iS4EventHeaderRow)
		.MergeCells = True
		.Font.Bold = True
		.Font.Color = $g_iTableMainHeaderFontColor
		.Font.Size = $g_iTableMainHeaderFontSize
		.Interior.Color = $g_iTableMainHeaderColorEvent
	EndWith

	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4EventIntroColumn) & $g_iS4EventLevelRow & ":" & _Excel_ColumnToLetter($g_iS4EventValueColumn - 1) & $g_iS4EventLevelRow)
		.MergeCells = True
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableMainHeaderColorEvent
	EndWith
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4EventIntroColumn) & $g_iS4EventDistanceRow & ":" & _Excel_ColumnToLetter($g_iS4EventValueColumn - 1) & $g_iS4EventDistanceRow)
		.MergeCells = True
		.Font.Bold = False
		.Font.Color = $g_iTableSecondHeaderFontColor
		.Interior.Color = $g_iTableMainHeaderColorEvent
	EndWith
	;Borders around input cells
	With $g_oWorkbook.Sheets($g_sNameOfSheet_4_GraphSelection).Range(_Excel_ColumnToLetter($g_iS4EventValueColumn) & $g_iS4EventLevelRow & ":" & _Excel_ColumnToLetter($g_iS4EventValueColumn) & $g_iS4EventDistanceRow )
		.WrapText = True
		.Borders.LineStyle = $g_xlContinuous
		.Borders.Weight = $xlThick
		.Borders.Color = $g_iTableMainHeaderColorEvent
	EndWith

	$g_oExcel.Application.ScreenUpdating = True ;Turn Screen Updating on
EndFunc   ;==>_Format_Sheet_4_GraphSelection


Func _Add_SelectionGraph_To_Sheet_3()
	$g_oExcel.Application.ScreenUpdating = False ;Turn Screen Updating off


	Local $l_sXValueRange, $l_asDataRange, $l_asDataName, $l_sTextContainter1, $l_sTextContainter2

	;Legger til utdragsgraf på side $g_sNameOfSheet_4_GraphSelection
	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), " ", "Lager utdragsgraf")
		$l_sChartHeading = 'Puls (Hjerteslag per minutt)'
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), " ", "Making graph selection ")
		$l_sChartHeading = 'Heart Rate (beats per minute)'
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1





	;Velger datapunkter fra $g_sNameOfSheet_5_CalculationsMainGraph
	$l_sXValueRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Column_TimeForShortGraph) & $g_iS5FirstContentRow & ":" & _Excel_ColumnToLetter($g_iS5Column_TimeForShortGraph) & $g_iS5LastSecondIntervalRow
	$l_asDataRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Column_PulseForShortGraph) & $g_iS5FirstContentRow & ":" & _Excel_ColumnToLetter($g_iS5Column_PulseForShortGraph) & $g_iS5LastSecondIntervalRow
	$l_asDataName = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Column_Pulse) & $g_iS5MainHeaderRow

	Local $l_iNumberOfColumnsToSpan = Round($g_iNumberOfSeconds / 50, 0)
	If $l_iNumberOfColumnsToSpan < 15 Then
		$l_iNumberOfColumnsToSpan = 15
	EndIf
	Local $l_iNumberOfColumnsToSpan = 25
	$l_sChartAreaRange = _Excel_ColumnToLetter($g_iS4GraphFirstColumn) & $g_iS4GraphFirstRow & ":" & _Excel_ColumnToLetter($g_iS4GraphFirstColumn+$l_iNumberOfColumnsToSpan) & $g_iS4GraphFirstRow+$l_iNumberOfColumnsToSpan
	;_XLChart_ChartCreate($oExcel, $vWorksheet, $iChartType, $sSizeByCells, $sChartName, $sXValueRange, $vDataRange, $vDataName[, $bShowLegend = True[, $sTitle = ""[, $sXTitle = ""[, $sYTitle = ""[, $sZTitle = ""[, $bShowDatatable = False[, $bScreenUpdate = False]]]]]]])
	Local $l_oChart = _XLChart_ChartCreate($g_oExcel, $g_sNameOfSheet_4_GraphSelection, $xlXYScatterLines,  $l_sChartAreaRange, $l_sChartHeading, $l_sXValueRange, $l_asDataRange, $l_asDataName, True, $l_sChartHeading)
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "Excel: _ChartCreate", "Error creating chart")
	EndIf

	;Find time of first datapoint
	Local $l_iFirstTime = $g_iHourMin + $g_iMinuteMin / 60
	;$l_iFirstTime = $l_iFirstTime - 1 / 60 ; Set first point 1 minute before first trackpoint

	Local $l_iLastTime = $g_iS4SelectedTo_Hour + $g_iS4SelectedTo_Minute / 60
	$l_iLastTime = $l_iLastTime + 1 / 60 ; Set last point 1 minute after last trackpoint

	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Justerer x-akse", "Lager utdragsgraf ")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Adjusting x-axis", "Making graph selection ")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1

	_XLChart_AxisSet($l_oChart.Axes(1), $l_iFirstTime / 24, $l_iLastTime / 24)
	With $l_oChart.Legend
		.Position = $xlLeft
		.IncludeInLayout = True
	EndWith
	With $l_oChart.Axes(1)
		If ($l_iLastTime - $l_iFirstTime) < 1 Then
			.MajorUnit = 1 / (24 * 6)
			.MinorUnit = 1 / (24 * 60 * 60)
		Else
			.MajorUnit = 1 / (24)
			.MinorUnit = 1 / (24 * 6)
		EndIf
	EndWith
	$g_oExcel.Application.ScreenUpdating = True ;Turn Screen Updating back on

	;Legge til Puls-gjennomsnitt, -standardavvik og -2*standardavvik
	Local $l_iStatisticsSeries
	;Gjennomsnitt
	$l_asDataName = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header) & $g_iS5Row_ShortGraphAverage_Header
	$l_sXValueRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header)& $g_iS5Row_ShortGraphAverage_FirstValue &":"& _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header)& $g_iS5Row_ShortGraphAverage_LastValue
	$l_asDataRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value)& $g_iS5Row_ShortGraphAverage_FirstValue &":"& _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value)& $g_iS5Row_ShortGraphAverage_LastValue

	$l_iStatisticsSeries = _XLChart_SeriesAdd($l_oChart, $l_sXValueRange, $l_asDataRange, $l_asDataName)
	With $l_iStatisticsSeries
		.Smooth = False
		.Format.Line.Visible = True
		.Format.Line.ForeColor.RGB = $COLOR_BLACK
		.MarkerStyle = -4142 ; xlMarkerStyleNone =	-4142
		.HasDataLabels = False
	EndWith

	;Gjennomsnitt pluss standardavvik
	$l_asDataName = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header) & $g_iS5Row_ShortGraphOneStDev_Header
	$l_sXValueRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header)& $g_iS5Row_ShortGraphOneStDev_FirstValue &":"& _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header)& $g_iS5Row_ShortGraphOneStDev_LastValue
	$l_asDataRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value)& $g_iS5Row_ShortGraphOneStDev_FirstValue &":"& _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value)& $g_iS5Row_ShortGraphOneStDev_LastValue

	$l_iStatisticsSeries = _XLChart_SeriesAdd($l_oChart, $l_sXValueRange, $l_asDataRange, $l_asDataName)
	With $l_iStatisticsSeries
		.Smooth = False
		.Format.Line.Visible = True
		.Format.Line.ForeColor.RGB = $COLOR_FUCHSIA
		.MarkerStyle = -4142 ; xlMarkerStyleNone =	-4142
		.HasDataLabels = False
	EndWith

	;Gjennomsnitt pluss to standardavvik
	$l_asDataName = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header) & $g_iS5Row_ShortGraphTwoStDev_Header
	$l_sXValueRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header)& $g_iS5Row_ShortGraphTwoStDev_FirstValue &":"& _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header)& $g_iS5Row_ShortGraphTwoStDev_LastValue
	$l_asDataRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value)& $g_iS5Row_ShortGraphTwoStDev_FirstValue &":"& _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value)& $g_iS5Row_ShortGraphTwoStDev_LastValue

	$l_iStatisticsSeries = _XLChart_SeriesAdd($l_oChart, $l_sXValueRange, $l_asDataRange, $l_asDataName)
	With $l_iStatisticsSeries
		.Smooth = False
		.Format.Line.Visible = True
		.Format.Line.ForeColor.RGB = $COLOR_RED
		.MarkerStyle = -4142 ; xlMarkerStyleNone =	-4142
		.HasDataLabels = False
	EndWith


	;_XLChart_LegendSet($l_oChart, Default, $xlLeft,Default, Default, Default, Default, Default)
	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Legger inn Aktivitetsmarkeringer", "Lager utdragsgraf ")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Adding Activity marks", "Making graph selection ")
	EndIf

	$g_iProgressCounter = $g_iProgressCounter + 1
	$g_oExcel.Application.ScreenUpdating = False ;Turn Screen Updating off

	;Legger til Aktivitetsmerker
	Local $l_oActivity_series, $l_iNumberOfVisibleActivities
	$l_iNumberOfVisibleActivities = 5


	;Local $l_sXActivityValueRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5FirstContentRow & ":" & _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5LastSecondIntervalRow
	Local $l_sXActivityValueRange
	Local $l_vActivityDataRange
	Local $l_vActivityNameRange

	$l_iCntr = 0
	While $l_iCntr < $g_iNumberOfActivities

		If $g_sChosenLanguage = "Norsk" Then
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Legger inn Aktivitetsmarkeringer: " & $l_iCntr + 1 & "/" & $g_iNumberOfActivities, "Lager utdragsgraf ")
		Else
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Adding Activity marks: " & $l_iCntr + 1 & "/" & $g_iNumberOfActivities, "Making graph selection ")
		EndIf
		$g_iProgressCounter = $g_iProgressCounter + 1


		$l_vActivityNameRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Activity_ForShort_ValueColumn) & $g_iS5Activity_ForShort_FirstHeaderRow + ($l_iCntr*$g_iS5Activity_ForShort_NumberOfRowsInSubTable)
		$l_sXActivityValueRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Activity_ForShort_TimeColumn)& $g_iS5Activity_ForShort_FirstContentRow +($l_iCntr*$g_iS5Activity_ForShort_NumberOfRowsInSubTable) &":"& _Excel_ColumnToLetter($g_iS5Activity_ForShort_TimeColumn)& $g_iS5Activity_ForShort_FirstContentRow +($l_iCntr*$g_iS5Activity_ForShort_NumberOfRowsInSubTable) +1
		$l_vActivityDataRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Activity_ForShort_ValueColumn)& $g_iS5Activity_ForShort_FirstContentRow +($l_iCntr*$g_iS5Activity_ForShort_NumberOfRowsInSubTable) &":"& _Excel_ColumnToLetter($g_iS5Activity_ForShort_ValueColumn)& $g_iS5Activity_ForShort_FirstContentRow +($l_iCntr*$g_iS5Activity_ForShort_NumberOfRowsInSubTable) +1


		$l_oActivity_series = _XLChart_SeriesAdd($l_oChart, $l_sXActivityValueRange, $l_vActivityDataRange, $l_vActivityNameRange)

		With $l_oActivity_series
			.Smooth = False
			.Format.Line.Visible = True
			.Format.Line.Weight = 3.5
			.MarkerStyle = 1 ; 1= xlMarkerStyleSquare
			.HasDataLabels = False
			If $l_iCntr >= $l_iNumberOfVisibleActivities Then
				.IsFiltered = True
			EndIf
		EndWith
		$l_iCntr = $l_iCntr + 1
	WEnd
	$g_oExcel.Application.ScreenUpdating = True ;Turn Screen Updating back on


	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Legger inn Hendelsesmarkeringer", "Lager utdragsgraf ")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Adding Event marks", "Making graph selection ")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	;Legger til Hendelsessmerker
	Local $l_oEvent_series, $l_iNumberOfVisibleEvents
	$l_iNumberOfVisibleEvents = 5
	$g_oExcel.Application.ScreenUpdating = False ;Turn Screen Updating off
	$l_iCntr = 0

	;Local $l_sXEventValueRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5FirstContentRow & ":" & _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5LastSecondIntervalRow
	Local $l_sXEventValueRange
	Local $l_vEventDataRange
	Local $l_vEventNameRange

	While $l_iCntr < $g_iNumberOfEvents
		If $g_sChosenLanguage = "Norsk" Then
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Legger inn Hendelsesmarkeringer: " & $l_iCntr + 1 & "/" & $g_iNumberOfEvents, "Lager utdragsgraf ")
		Else
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Adding Event marks: " & $l_iCntr + 1 & "/" & $g_iNumberOfEvents, "Making graph selection ")
		EndIf
		$g_iProgressCounter = $g_iProgressCounter + 1


		$l_vEventNameRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" & _Excel_ColumnToLetter($g_iS5Event_ForShort_ValueColumn) & $g_iS5Event_ForShort_FirstHeaderRow + ($l_iCntr*$g_iS5Event_ForShort_NumberOfRowsInSubTable)
		$l_sXEventValueRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Event_ForShort_TimeColumn)& $g_iS5Event_ForShort_FirstContentRow +($l_iCntr*$g_iS5Event_ForShort_NumberOfRowsInSubTable) &":"& _Excel_ColumnToLetter($g_iS5Event_ForShort_TimeColumn)& $g_iS5Event_ForShort_FirstContentRow +($l_iCntr*$g_iS5Event_ForShort_NumberOfRowsInSubTable) +1
		$l_vEventDataRange = "='" & $g_sNameOfSheet_5_CalculationsMainGraph & "'!" &_Excel_ColumnToLetter($g_iS5Event_ForShort_ValueColumn)& $g_iS5Event_ForShort_FirstContentRow +($l_iCntr*$g_iS5Event_ForShort_NumberOfRowsInSubTable) &":"& _Excel_ColumnToLetter($g_iS5Event_ForShort_ValueColumn)& $g_iS5Event_ForShort_FirstContentRow +($l_iCntr*$g_iS5Event_ForShort_NumberOfRowsInSubTable) +1

		$l_oEvent_series = _XLChart_SeriesAdd($l_oChart, $l_sXEventValueRange, $l_vEventDataRange, $l_vEventNameRange)
		; "='"&$g_sNameOfSheet_5_CalculationsMainGraph&"'!R5C"&$l_iStartColumnEvent+$l_iCntr&":R"&$l_iLastRow&"C"&$l_iStartColumnEvent+$l_iCntr, "='"&$g_sNameOfSheet_5_CalculationsMainGraph&"'!R4C"&$l_iStartColumnEvent+$l_iCntr)
		With $l_oEvent_series
			.Smooth = False
			.Format.Line.Visible = False
			.MarkerStyle = 8 ; xlMarkerStyleCircle	8	Circular markers
			.MarkerSize = 10 ; min = 2, max = 72
			.HasDataLabels = True
			.DataLabels.ShowValue = False
			.DataLabels.ShowSeriesName = True
			.DataLabels.Position = 0 ;xlLabelPositionAbove	 = 0	Data label is positioned above the data point.
			.DataLabels.Orientation = -4171 ; xlUpward= 	-4171	Text runs upward.
			If $l_iCntr >= $l_iNumberOfVisibleEvents Then
				.IsFiltered = True
			EndIf
		EndWith

		$l_iCntr = $l_iCntr + 1
	WEnd

	$g_oExcel.Application.ScreenUpdating = True ;Turn Screen Updating on
EndFunc   ;==>_Add_SelectionGraph_To_Sheet_3

Func OnCreateExcel()
	;Creating Com error handler
	Local $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")

	If $g_sChosenLanguage = "Norsk" Then
		ProgressOn("Forstå meg: Framdrift", "Omgjør rådata til XML", "Venter på GPSBabel", -2, -1, $DLG_MOVEABLE)
	Else
		ProgressOn("Understand How I Am: Progress", "Transforming raw data to XML", "Waiting for GPSBabel", -2, -1, $DLG_MOVEABLE)
	EndIf
	Local $l_iResult, $l_iCntr

	$l_iResult = _ControlFolderPaths()
	Local $l_iErrorInt
	If @error Then
		$l_iErrorInt = @error
		;MsgBox(0, "FEIL", "FEIL. @Error = " &$l_iErrorInt)
		If ($l_iErrorInt = 1 Or $l_iErrorInt = 2 Or $l_iErrorInt = 3) Then
			ProgressOff()
			$g_iProgressCounter = 0
			_SetConfiguration(1)
		Else
			ProgressOff()
			$g_iProgressCounter = 0
			_SetConfiguration(2)
		EndIf
		Return SetError($l_iResult, @error, 0)
	EndIf


 	$l_iResult = _TransformToXML()
 	If @error Then
 		ProgressOff()
		$g_iProgressCounter = 0
 		Return SetError($l_iResult, @error, 0)
 	EndIf

	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet (100*(Round(($g_iProgressCounter +1) /$g_iProgressMaxCount,2)) , "Fra Garmin .fit fil til .xml fil" , "Omgjør data")
	Else
		ProgressSet (100*(Round(($g_iProgressCounter +1) /$g_iProgressMaxCount,2)) , "From Garmin .fit file to .xml fil" , "Transforming data")
	EndIf
 	$g_iProgressCounter = $g_iProgressCounter +1
 	Sleep(2000)

	Local $l_sSourceFileXML = $g_sIniValue_XMLFolder & "\" & $g_sCurrentFileName & ".xml"
	Local $l_sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
	Local $l_aPathSplit = _PathSplit($l_sSourceFileXML, $l_sDrive, $sDir, $sFileName, $sExtension)
	Local $l_sXLSX = $g_sIniValue_ExcelFolder & "\" & $g_sCurrentFileName & ".xlsx"

	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "", "Åpner Excel")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "", "Opening Excel")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	$g_oExcel = _Excel_Open(True)

	If @error Then
		If $g_sChosenLanguage = "Norsk" Then
			MsgBox($MB_SYSTEMMODAL, "Excel UDF: Åpne Excel", "FEIL ved opprettelse av EXCEL object" & @CRLF & " @error = " & @error & ", @extended = " & @extended)
		Else
			MsgBox($MB_SYSTEMMODAL, "Excel UDF: Open Excel", "ERROR creating EXCEL object" & @CRLF & " @error = " & @error & ", @extended = " & @extended)
		EndIf
	Else
		If $g_sChosenLanguage = "Norsk" Then
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Suksess", "Åpner Excel")
		Else
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Success", "Opening Excel")
		EndIf
		$g_iProgressCounter = $g_iProgressCounter + 1
	EndIf
	Sleep(5000)

	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Henter inn målinger", "Oppretter Excel arbeidsbok")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Collecting registered data", "Creating Excel workbook")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	Local $xlXmlLoadImportToList = 2 ; Places the contents of the XML data file in an XML table
	$g_oWorkbook = $g_oExcel.Workbooks.OpenXML($l_sSourceFileXML, Default, $xlXmlLoadImportToList)
	If @error Then
		If $g_sChosenLanguage = "Norsk" Then
			MsgBox($MB_SYSTEMMODAL, "Excel UDF: Opprette arbeidsbok Excel", "FEIL ved opprettelse av EXCEL arbeidsbok" & @CRLF & " @error = " & @error & ", @extended = " & @extended)
		Else
			MsgBox($MB_SYSTEMMODAL, "Excel UDF: Create Excel workbook", "ERROR creating EXCEL workbook" & @CRLF & " @error = " & @error & ", @extended = " & @extended)
		EndIf
	Else
		If $g_sChosenLanguage = "Norsk" Then
			ProgressSet(100 * (Round($g_iProgressCounter / $g_iProgressMaxCount, 2)), "Suksess", "Oppretter Excel arbeidsbok")
		Else
			ProgressSet(100 * (Round($g_iProgressCounter / $g_iProgressMaxCount, 2)), "Success",  "Creating Excel workbook")
		EndIf
		$g_iProgressCounter = $g_iProgressCounter + 1
	EndIf
	Sleep(2000)
	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Navn til arbeidsbok:" & @CRLF & $g_sCurrentFileName & ".xlsx", "Lagrer Excel arbeidsbok")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Name of workbook:" & @CRLF & $g_sCurrentFileName & ".xlsx", "Saving Excel workbook")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1

	$l_iResult = _Excel_BookSaveAs($g_oWorkbook, $l_sXLSX, $xlWorkbookDefault, True)

	If @error Then
		If $g_sChosenLanguage = "Norsk" Then
			MsgBox($MB_SYSTEMMODAL, "Excel UDF: Lagre excelfil", "FEIL under lagring av arbeidsbok. Filnavn:" & @CRLF  & $l_sXLSX & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		Else
			MsgBox($MB_SYSTEMMODAL, "Excel UDF: Save Excel file", "ERROR saving workbook. Fils name:" & @CRLF  & $l_sXLSX & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		EndIf
	Else
		If $g_sChosenLanguage = "Norsk" Then
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Navn til arbeidsbok: "  & @CRLF & $g_sCurrentFileName & ".xlsx"& @CRLF &"Suksess", "Lagrer Excel arbeidsbok")
		Else
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Name of workbook: "  & @CRLF & $g_sCurrentFileName & ".xlsx"& @CRLF &"Success", "Saving Excel workbook")
		EndIf
		$g_iProgressCounter = $g_iProgressCounter + 1
	EndIf

	;TODO - sjekker antall rader
	Local $l_iUsedRowsCount = $g_oWorkbook.Sheets(1).UsedRange.Rows.Count
	If ($l_iUsedRowsCount < 15) Then
		If $g_sChosenLanguage = "Norsk" Then
			MsgBox($MB_SYSTEMMODAL, "Ikke nok datapunkter", "Filen har for få datapunkter. Avslutter.")
		Else
			MsgBox($MB_SYSTEMMODAL, "Not enough data points", "The file does not have enough data points. Terminating.")
		EndIf

		ProgressOff()
		$g_iProgressCounter = 0
		Return 0
	EndIf

	Local $l_iUsedColumnsCount = $g_oWorkbook.Sheets(1).UsedRange.Columns.Count
	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Navn til arbeidsbok: "  & @CRLF & $g_sCurrentFileName & ".xlsx"& @CRLF &" ", "Finner kolonner for tid og puls")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Name of workbook: "  & @CRLF & $g_sCurrentFileName & ".xlsx"& @CRLF &" ", "Finding columns for time and heart rate")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1


	Local $l_sColumnOfDateTime = "L"
	Local $l_sColumnOfPulse = "M"
	Local $l_iColumnOfDateTime, $l_iColumnOfPulse;  = _Excel_ColumnToNumber($l_sColumnOfDateTime)
;	Local $l_iColumnOfPulse = _Excel_ColumnToNumber($l_sColumnOfPulse)


	;Checking Date/time column:  L ns1:Time, M ns1:Value5
	Local $l_sContentOfColumnL = _Excel_RangeRead($g_oWorkbook, 1, "L1")
	$l_iColumnOfDateTime = _Excel_ColumnToNumber($l_sColumnOfDateTime)

	Local $l_sContentOfColumnM = _Excel_RangeRead($g_oWorkbook, 1, "M1")
	Local $l_sContentOfColumnK = _Excel_RangeRead($g_oWorkbook, 1, "K1")


	;MsgBox($MB_SYSTEMMODAL, "Checking column L header","Content of column L: " & $l_sContentOfColumnL)
	If StringCompare($l_sContentOfColumnL, "ns1:Time") =0 Then
		If $g_sChosenLanguage = "Norsk" Then
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Navn til arbeidsbok: " & @CRLF & $g_sCurrentFileName & ".xlsx"& @CRLF &" ", "Fant tid i kolonne L")
		Else
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Name of workbook: "  & @CRLF & $g_sCurrentFileName & ".xlsx"& @CRLF &" ", "Found time in column L")
		EndIf
		$g_iProgressCounter = $g_iProgressCounter + 1
		;MsgBox($MB_SYSTEMMODAL, "Success", "Found Time content in column L " )
		$l_sColumnOfDateTime = "L"
		$l_iColumnOfDateTime = _Excel_ColumnToNumber($l_sColumnOfDateTime)
	ElseIf  StringCompare($l_sContentOfColumnK, "ns1:Time") =0 Then
		If $g_sChosenLanguage = "Norsk" Then
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Navn til arbeidsbok: " & @CRLF & $g_sCurrentFileName & ".xlsx"& @CRLF &" ", "Fant TID i kolonne K")
		Else
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Name of workbook: "  & @CRLF & $g_sCurrentFileName & ".xlsx"& @CRLF &" ", "Found TIME in column K")
		EndIf
		$g_iProgressCounter = $g_iProgressCounter + 1
		;MsgBox($MB_SYSTEMMODAL, "Success", "Found Time content in column K " )
		$l_sColumnOfDateTime = "K"
		$l_iColumnOfDateTime = _Excel_ColumnToNumber($l_sColumnOfDateTime)
	Else
		If $g_sChosenLanguage = "Norsk" Then
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Navn til arbeidsbok: " & @CRLF & $g_sCurrentFileName &	".xlsx"& @CRLF &" ", "Venter på manuell input av TIDskolonne")
			$g_iProgressCounter = $g_iProgressCounter + 1
			$l_sColumnOfDateTime = InputBox("DATO / TID", "Skriv inn kolonnen der DATO/TID befinner seg", "", " M3")
		Else
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Name of workbook: "  & @CRLF & $g_sCurrentFileName & ".xlsx"& @CRLF &" ", "Waiting for manual input: Column letter for TIME column")
			$g_iProgressCounter = $g_iProgressCounter + 1
			$l_sColumnOfDateTime = InputBox("DATE / TIME", "Write the letter for the column where DATE/TIME is found", "", " M3")
		EndIf
	EndIf

	$l_iColumnOfPulse = $l_iColumnOfDateTime +1
	$l_sColumnOfPulse = _Excel_ColumnToLetter($l_iColumnOfPulse)
	;MsgBox($MB_SYSTEMMODAL, "Checking for PULSE", "Checking column " & $l_sColumnOfPulse & " that has number " & $l_iColumnOfPulse)

	Local $l_sContentOfColumnAfterTime = _Excel_RangeRead($g_oWorkbook, 1, $l_sColumnOfPulse &"1")


	If StringCompare($l_sContentOfColumnAfterTime, "ns1:Value5") =0 Then
		If $g_sChosenLanguage = "Norsk" Then
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Navn til arbeidsbok: " & @CRLF & $g_sCurrentFileName & ".xlsx"& @CRLF &" ", "Fant PULS i kolonne "& $l_sColumnOfPulse)
		Else
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Name of workbook: " & @CRLF & $g_sCurrentFileName & ".xlsx"& @CRLF &" ", "Found HEART RATE in column "& $l_sColumnOfPulse)
		EndIf
		$g_iProgressCounter = $g_iProgressCounter + 1
		;MsgBox($MB_SYSTEMMODAL, "Success", "Found PULSE content in column " &$l_sColumnOfPulse )
	Else
		If $g_sChosenLanguage = "Norsk" Then
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Navn til arbeidsbok: " & @CRLF & $g_sCurrentFileName & ".xlsx"& @CRLF & " ", "Venter på manuell input av PULSkolonne")
			$g_iProgressCounter = $g_iProgressCounter + 1
			$l_sColumnOfPulse = InputBox("PULS", "Skriv inn kolonnen der PULS befinner seg", "", " M3") ;, -1, -1, 0,0)
		Else
			ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Name of workbook: " & @CRLF & $g_sCurrentFileName & ".xlsx"& @CRLF &" ", "Waiting for manual input: Column letter for HEART RATE column")
			$g_iProgressCounter = $g_iProgressCounter + 1
			$l_sColumnOfPulse = InputBox("HEART RATE", "Write the letter for the column where HEART RATE is found", "", " M3") ;, -1, -1, 0,0)
		EndIf
	EndIf

	$l_iColumnOfDateTime = _Excel_ColumnToNumber($l_sColumnOfDateTime)
	$l_iColumnOfPulse = _Excel_ColumnToNumber($l_sColumnOfPulse)


	Local $l_iFirstColumnToRemove
	If $l_iColumnOfPulse > $l_iColumnOfDateTime Then
		$l_iColumnsBetween = $l_iColumnOfPulse - $l_iColumnOfDateTime
		If ($l_iColumnsBetween) > 1 Then
			_Excel_RangeDelete($g_oWorkbook.Worksheets(1), _Excel_ColumnToLetter($l_iColumnOfDateTime + 1) & ":" & _Excel_ColumnToLetter($l_iColumnOfPulse - 1))
			$l_iColumnOfPulse = $l_iColumnOfDateTime + 1
			$l_sColumnOfPulse = _Excel_ColumnToLetter($l_iColumnOfPulse)
		EndIf

		$l_iFirstColumnToRemove = $l_iColumnOfPulse + 1
		If $l_iUsedColumnsCount > $l_iFirstColumnToRemove Then
			_Excel_RangeDelete($g_oWorkbook.Worksheets(1), _Excel_ColumnToLetter($l_iFirstColumnToRemove) & ":" & _Excel_ColumnToLetter($l_iUsedColumnsCount))
		EndIf

		If $l_iColumnOfDateTime > 1 Then
			_Excel_RangeDelete($g_oWorkbook.Worksheets(1), "A:" & _Excel_ColumnToLetter($l_iColumnOfDateTime - 1))
		EndIf

	Else
		If $g_sChosenLanguage = "Norsk" Then
			MsgBox(0, "Feil rekkefølge", "PULS KOMMER FØR DATO TID!! Avslutter omforming")
		Else
			MsgBox(0, "Wrong order", "HEART RATE PLACED BEFORE DATE/TIME!! Terminating transforma")
		EndIf

		ProgressOff()
		$g_iProgressCounter = 0
		Return 0
	EndIf


	;Add sheets in front of sheet with data
	_Excel_SheetAdd($g_oWorkbook, 1, True, $g_iNumberOfSheets-1)

	;Name each sheet
	If $g_iNumberOfSheets = 5 Then
		$g_oWorkbook.Sheets(1).Name = $g_sNameOfSheet_1_Descriptions
		$g_oWorkbook.Sheets(2).Name = $g_sNameOfSheet_2_Observations
		$g_oWorkbook.Sheets(3).Name = $g_sNameOfSheet_3_GraphMain
		$g_oWorkbook.Sheets(4).Name = $g_sNameOfSheet_4_GraphSelection
		$g_oWorkbook.Sheets(5).Name = $g_sNameOfSheet_5_CalculationsMainGraph
	ElseIf $g_iNumberOfSheets = 4 Then
		$g_oWorkbook.Sheets(1).Name = $g_sNameOfSheet_2_Observations
		$g_oWorkbook.Sheets(2).Name = $g_sNameOfSheet_3_GraphMain
		$g_oWorkbook.Sheets(3).Name = $g_sNameOfSheet_4_GraphSelection
		$g_oWorkbook.Sheets(4).Name = $g_sNameOfSheet_5_CalculationsMainGraph
	Else
		$g_oWorkbook.Sheets(1).Name = $g_sNameOfSheet_2_Observations
		$g_oWorkbook.Sheets(2).Name = $g_sNameOfSheet_3_GraphMain
		$g_oWorkbook.Sheets(3).Name = $g_sNameOfSheet_5_CalculationsMainGraph

	EndIf

	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Splitter tid og dato", "Rydder i Excel-fil")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Splitting time and date", "Rearranging Excel file")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1


	$g_iS5FirstContentRow = 2
	; Insert 1 columns before colum Pulse on the active worksheet
	_Excel_RangeInsert($g_oWorkbook.Worksheets($g_sNameOfSheet_5_CalculationsMainGraph), "B:B", Default, Default)
	;MsgBox(0, "La inn i mellom", "")
	;_Excel_RangeInsert($g_oWorkbook.Worksheets(1), "B")
	With $g_oWorkbook.Sheets($g_sNameOfSheet_5_CalculationsMainGraph).Range("A:A")
		.TextToColumns(Default, Default, Default, Default, False, False, False, False, True, "T", Default, Default, Default, Default)
	EndWith
	; Insert 1 columns before colum Pulse on the active worksheet
	_Excel_RangeInsert($g_oWorkbook.Worksheets($g_sNameOfSheet_5_CalculationsMainGraph), "C:E", Default, Default)
	With $g_oWorkbook.Sheets($g_sNameOfSheet_5_CalculationsMainGraph).Range("B:B")
		.TextToColumns(Default, Default, Default, Default, False, False, False, False, True, ":", Default, Default, Default, Default)
	EndWith
	With $g_oWorkbook.Sheets($g_sNameOfSheet_5_CalculationsMainGraph).Range("D:D")
		.TextToColumns(Default, Default, Default, Default, False, False, False, False, True, "Z", Default, Default, Default, Default)
	EndWith

	If $g_sChosenLanguage = "Norsk" Then
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, "Dato", "A1")
		; Insert 1 columns before colum Minutt on the active worksheet
		_Excel_RangeInsert($g_oWorkbook.Worksheets($g_sNameOfSheet_5_CalculationsMainGraph), "C:C", Default, Default)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, "=B" & $g_iS5FirstContentRow & "+" & "'" & $g_sNameOfSheet_2_Observations & "'!$" & _Excel_ColumnToLetter($g_iS2HourCorrectionInputColumn) & "$" & $g_iS2HourCorrectionRow, "C" & $g_iS5FirstContentRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, "Time", "B1")
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, "Korrigert Time", "C1")
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, "Minutt", "D1")
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, "Sekund", "E1")
	Else
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, "Date", "A1")
		; Insert 1 columns before colum Minutt on the active worksheet
		_Excel_RangeInsert($g_oWorkbook.Worksheets($g_sNameOfSheet_5_CalculationsMainGraph), "C:C", Default, Default)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, "=B" & $g_iS5FirstContentRow & "+" & "'" & $g_sNameOfSheet_2_Observations & "'!$" & _Excel_ColumnToLetter($g_iS2HourCorrectionInputColumn) & "$" & $g_iS2HourCorrectionRow, "C" & $g_iS5FirstContentRow)
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, "Hour", "B1")
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, "Corrected Hour", "C1")
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, "Minute", "D1")
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, "Seconds", "E1")
	EndIf
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $g_sS5CorrectedTimeForMainTable_ColumnName, "F1")
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $g_sS5PulseForMainTable_ColumnName, "G1")


	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Korrigerer tid", "Rydder i Excel-fil")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Correcting time", "Rearranging Excel file")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1

	$g_oWorkbook.Sheets($g_sNameOfSheet_5_CalculationsMainGraph).Range("B:E").NumberFormat = "0"

	;Setting name on Table
	$g_oWorkbook.Sheets($g_sNameOfSheet_5_CalculationsMainGraph).ListObjects(1).Name = $g_sS5MainTable_Name

	Local $l_bDeleteTopRows = True
	Local $l_iLastRowToDelete = -1

	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Fjerner tomme rader", "Rydder i Excel-fil")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Removing empty rows", "Rearranging Excel file")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1

	While $l_bDeleteTopRows
		_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, "=MATCH(TRUE;INDEX(G" & $g_iS5FirstContentRow & ":G" & $g_iS5FirstContentRow + 99 & "<>0;);0)", "K" & $g_iS5FirstContentRow)

		; MsgBox(0,"klar til sletting", "Satte in formel for å finne første rad:   (G2:G11<>0;);")
		$l_iLastRowToDelete = _Excel_RangeRead($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, "K" & $g_iS5FirstContentRow)
		If $l_iLastRowToDelete == 1 Then
			;MsgBox(0,"klar til sletting", "Ingen rader å slette")
			$l_bDeleteTopRows = False
			ExitLoop

		ElseIf $l_iLastRowToDelete > 1 Then
			;MsgBox(0,"klar til sletting", "Første rad: 2" &@CRLF&"Siste rad"&$l_iLastRowToDelete)
			_Excel_RangeDelete($g_oWorkbook.Worksheets($g_sNameOfSheet_5_CalculationsMainGraph), $g_iS5FirstContentRow & ":" & $l_iLastRowToDelete)

		Else
			;MsgBox(0,"sletter rad", "sletter øverste rader 2:101")
			_Excel_RangeDelete($g_oWorkbook.Worksheets($g_sNameOfSheet_5_CalculationsMainGraph), $g_iS5FirstContentRow & ":" & $g_iS5FirstContentRow + 99)
		EndIf

	WEnd

	_Excel_RangeDelete($g_oWorkbook.Worksheets($g_sNameOfSheet_5_CalculationsMainGraph), "K" & $g_iS5FirstContentRow)



	;Find number of trackpoints
	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Antall datapunkter: ", "Beregner datapunkter")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Number of data points: ", "Calculating data points")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	Local $xlup = -4162
	Local $l_iNumberOfTrackpoints = $g_oWorkbook.Worksheets($g_sNameOfSheet_5_CalculationsMainGraph).Range("B65536").End($xlup).Row - 1

	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Antall datapunkter: " & $l_iNumberOfTrackpoints, "Beregner datapunkter")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Number of data points:" & $l_iNumberOfTrackpoints, "Calculating data points")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1


	$g_iS5LastContentRow = $g_iS5FirstContentRow + $l_iNumberOfTrackpoints - 1

	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, "00:00:00", _Excel_ColumnToLetter($g_iS5Column_CorrectedTime) & $g_iS5FirstContentRow)

	Local $l_sFormula = "=TIME(" & _Excel_ColumnToLetter($g_iS5Column_CorrectedHour) & $g_iS5FirstContentRow & ";" & _Excel_ColumnToLetter($g_iS5Column_Minute) & $g_iS5FirstContentRow & ";" & _Excel_ColumnToLetter($g_iS5Column_Second) & $g_iS5FirstContentRow & ")"
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_CorrectedTime) & $g_iS5FirstContentRow)

	$g_iHourMin = _Excel_RangeRead($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, _Excel_ColumnToLetter($g_iS5Column_CorrectedHour) & $g_iS5FirstContentRow, 1) ;Hour
	$g_iMinuteMin = _Excel_RangeRead($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, _Excel_ColumnToLetter($g_iS5Column_Minute) & $g_iS5FirstContentRow, 1) ;Minute
	$g_iSecondMin = _Excel_RangeRead($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, _Excel_ColumnToLetter($g_iS5Column_Second) & $g_iS5FirstContentRow, 1) ;Second

	$g_iHourMax = _Excel_RangeRead($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, _Excel_ColumnToLetter($g_iS5Column_CorrectedHour) & $g_iS5LastContentRow, 1) ;Hour
	$g_iMinuteMax = _Excel_RangeRead($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, _Excel_ColumnToLetter($g_iS5Column_Minute) & $g_iS5LastContentRow, 1) ;Minute
	$g_iSecondMax = _Excel_RangeRead($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, _Excel_ColumnToLetter($g_iS5Column_Second) & $g_iS5LastContentRow, 1) ;Second

	;MsgBox(0,"", "Første tid: "&$g_iHourMin&":"&$g_iMinuteMin&":"&$g_iSecondMin&@CRLF & "Siste tid: "&$g_iHourMax&":"&$g_iMinuteMax&":"&$g_iSecondMax)
	$g_iNumberOfMinutes = ($g_iHourMax - $g_iHourMin) * 60 + ($g_iMinuteMax - $g_iMinuteMin) + 1
	$g_iNumberOfSeconds = Round((($g_iHourMax - $g_iHourMin) * 60 * 60 + ($g_iMinuteMax - $g_iMinuteMin) * 60 + ($g_iSecondMax - $g_iSecondMin) + 1) / $g_iNumberOfSecondsInSecondInterval)
	;MsgBox(0,"", "Antall sekunder: "&$g_iNumberOfSeconds)

	$g_iS5LastMinuteIntervalRow = $g_iS5FirstContentRow + $g_iNumberOfMinutes - 1
	$g_iS5LastSecondIntervalRow = $g_iS5FirstContentRow + $g_iNumberOfSeconds - 1
	Local $l_iNumberOfExtraRowsForSeconds = $g_iS5LastSecondIntervalRow - $g_iS5LastContentRow

	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Lager sekundkolonne", "Oppdaterer tabell")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Making seconds column", "Updating table")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	;Creating Seconds table with autotable

	Local $l_sSecondsRange = _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5MainHeaderRow & ":" & _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5LastSecondIntervalRow
	Local $l_sSecondsTopRightRange = _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5MainHeaderRow
	;Local $l_xlSrcRange = 1

	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $g_sS5SecondInterval_ColumnName, _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5MainHeaderRow)
	;_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, "00:00:00", _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5FirstContentRow)

	;$l_sFormula = "=$" & _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & "$" & $g_iS5FirstContentRow & "+TIME(0;0;ROW(" & _Excel_ColumnToLetter($g_iS5Column_Hour) & $g_iS5FirstContentRow + 1 & ")-ROW($" & _Excel_ColumnToLetter($g_iS5Column_Hour) & "$" & $g_iS5FirstContentRow & "))"
	$l_sFormula = "=" & _Excel_ColumnToLetter($g_iS5Column_SecondInterval) &  $g_iS5FirstContentRow & "+TIME(0;0;1)"
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5FirstContentRow + 1)

	Local $l_sCopyRange = _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5FirstContentRow
	IF $g_iS5LastSecondIntervalRow > $g_iS5LastContentRow Then
		_Excel_RangeCopyPaste($g_oWorkbook.Worksheets($g_sNameOfSheet_5_CalculationsMainGraph), $l_sCopyRange, _
		$g_oWorkbook.Worksheets($g_sNameOfSheet_5_CalculationsMainGraph).Range(_Excel_ColumnToLetter($g_iS5Column_SecondInterval)&$g_iS5LastContentRow + 1 &":" & _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5LastSecondIntervalRow))
		;_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5LastContentRow + 1 &":" & _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5LastSecondIntervalRow)
	EndIf


	$g_oExcel.Application.AutoCorrect.AutoFillFormulasInLists = False
	$l_sFormula = "=TIME(" & _Excel_ColumnToLetter($g_iS5Column_CorrectedHour) & $g_iS5FirstContentRow & ";" & _Excel_ColumnToLetter($g_iS5Column_Minute) & $g_iS5FirstContentRow & ";" & _Excel_ColumnToLetter($g_iS5Column_Second) & $g_iS5FirstContentRow & ")"
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5FirstContentRow)

	$g_oExcel.Application.AutoCorrect.AutoFillFormulasInLists = True


	;Creating header for longGraphTable
	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Lager kolonner for hovedgraf", "Oppdaterer tabell")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Making columns for main graph", "Updating table")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $g_sS5TimeForLongGraph_ColumnName, _Excel_ColumnToLetter($g_iS5Column_TimeForLongGraph) & $g_iS5MainHeaderRow)
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $g_sS5PulseForLongGraph_ColumnName, _Excel_ColumnToLetter($g_iS5Column_PulseForLongGraph) & $g_iS5MainHeaderRow)

	;Creating header for shortGraphTable
	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Lager kolonner for kortgraf", "Oppdaterer tabell")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Making columns for short graph", "Updating table")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $g_sS5TimeForShortGraph_ColumnName, _Excel_ColumnToLetter($g_iS5Column_TimeForShortGraph) & $g_iS5MainHeaderRow)
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $g_sS5PulseForShortGraph_ColumnName, _Excel_ColumnToLetter($g_iS5Column_PulseForShortGraph) & $g_iS5MainHeaderRow)

	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Lager tabeller for beskrivelse "& @CR & "av aktiviteter og hendelser", "Lager ark " & $g_sNameOfSheet_1_Descriptions)
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Making tables for descriptions "& @CR & "of activities and events", "Making sheet " & $g_sNameOfSheet_1_Descriptions)
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	_Fill_Sheet_1_Descriptions()

	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Formaterer tabeller for beskrivelse "& @CR & "av aktiviteter og hendelser", "Lager ark " & $g_sNameOfSheet_1_Descriptions)
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Formatting tables for descriptions "& @CR & "of activities and events", "Making sheet " & $g_sNameOfSheet_1_Descriptions)
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	_Format_Sheet_1_Descriptions()


	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Lager tabeller for observasjoner", "Lager ark " & $g_sNameOfSheet_2_Observations)
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Making tables for observations", "Making sheet " & $g_sNameOfSheet_2_Observations)
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	_Fill_Sheet_2_Observations($l_iNumberOfTrackpoints)

	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Formaterer tabeller for observasjoner", "Formaterer ark " & $g_sNameOfSheet_2_Observations)
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Formatting tables for observations", "Making sheet " & $g_sNameOfSheet_2_Observations)
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	_Format_Sheet_2_Observations()


	;Create activityTable for long graph on sheet 5 for chart
	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Lager tabell for aktivitet i hovedgraf", "Lager tabell")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Making table for activities in main graph", "Making table")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1

	If $g_sChosenLanguage = "Norsk" Then
		$l_sTextContainter1 = "Hovedgraf"
	Else
		$l_sTextContainter1 = "Main graph"
	EndIf

	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sTextContainter1, _Excel_ColumnToLetter($g_iS5Activity_ForLong_HeaderColumn)&"1")
	$l_iCntr = 0
	Local $l_iActivityTableWidth = 3
	Local $l_iActivityTableHeight = $g_iNumberOfActivities*($g_iS5Activity_ForLong_NumberOfRowsInSubTable)
	Local $l_aActivityArray[$l_iActivityTableHeight][$l_iActivityTableWidth]
	Local $l_sFromTimeFormula, $l_sToTimeFormula, $l_sLevelFormula

	While $l_iCntr < $g_iNumberOfActivities
		$l_aActivityArray[$l_iCntr*$g_iS5Activity_ForLong_NumberOfRowsInSubTable +0][0] = "='"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2ActivityFirstColumn)&$g_iS2ActivitySecondHeaderRow+1+$l_iCntr
		$l_aActivityArray[$l_iCntr*$g_iS5Activity_ForLong_NumberOfRowsInSubTable +0][1] = "='"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2ActivityDescriptionColumn)&$g_iS2ActivitySecondHeaderRow+1+$l_iCntr
		$l_sFromTimeFormula = "='"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2ActivityFromColumn)&$g_iS2ActivitySecondHeaderRow+1+$l_iCntr
		$l_sToTimeFormula = "='"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2ActivityToColumn)&$g_iS2ActivitySecondHeaderRow+1+$l_iCntr
		$l_sLevelFormula = "='"&$g_sNameOfSheet_3_GraphMain &"'!"& _Excel_ColumnToLetter($g_iS3ActivityValueColumn)& $g_iS3ActivityLevelRow & "-" & _
			"'"&$g_sNameOfSheet_3_GraphMain &"'!"& _Excel_ColumnToLetter($g_iS3ActivityValueColumn)& $g_iS3ActivityDistanceRow & "+" & _
			"'"&$g_sNameOfSheet_3_GraphMain &"'!"& _Excel_ColumnToLetter($g_iS3ActivityValueColumn)& $g_iS3ActivityDistanceRow & "*" & "'"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2ActivityFirstColumn)&$g_iS2ActivitySecondHeaderRow+1+$l_iCntr
		$l_aActivityArray[$l_iCntr*$g_iS5Activity_ForLong_NumberOfRowsInSubTable +1][0] = $l_sFromTimeFormula
		$l_aActivityArray[$l_iCntr*$g_iS5Activity_ForLong_NumberOfRowsInSubTable +1][1] = $l_sLevelFormula

		$l_aActivityArray[$l_iCntr*$g_iS5Activity_ForLong_NumberOfRowsInSubTable +2][0] = $l_sToTimeFormula
		$l_aActivityArray[$l_iCntr*$g_iS5Activity_ForLong_NumberOfRowsInSubTable +2][1] = $l_sLevelFormula
		$l_iCntr = $l_iCntr + 1
	WEnd

	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_aActivityArray, _Excel_ColumnToLetter($g_iS5Activity_ForLong_TimeColumn) & $g_iS5Activity_ForLong_FirstHeaderRow)

	;Create eventTable for long graph on sheet 5 for chart
	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Lager tabell for hendelser i hovedgraf", "Lager tabell")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Making table for events in main graph", "Making table")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1

	$l_iCntr = 0
	Local $l_iEventTableWidth = 3
	Local $l_iEventTableHeight = $g_iNumberOfEvents*($g_iS5Event_ForLong_NumberOfRowsInSubTable)
	Local $l_aEventArray[$l_iEventTableHeight][$l_iEventTableWidth]

	While $l_iCntr < $g_iNumberOfEvents
		$l_aEventArray[$l_iCntr*$g_iS5Event_ForLong_NumberOfRowsInSubTable +0][0] = "='"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2EventFirstColumn)&$g_iS2EventSecondHeaderRow+1+$l_iCntr
		$l_aEventArray[$l_iCntr*$g_iS5Event_ForLong_NumberOfRowsInSubTable +0][1] = "='"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2EventDescriptionColumn)&$g_iS2EventSecondHeaderRow+1+$l_iCntr

		$l_sLevelFormula = "='"&$g_sNameOfSheet_3_GraphMain &"'!"& _Excel_ColumnToLetter($g_iS3EventValueColumn)& $g_iS3EventLevelRow & "-" & _
			"'"&$g_sNameOfSheet_3_GraphMain &"'!"& _Excel_ColumnToLetter($g_iS3EventValueColumn)& $g_iS3EventDistanceRow & "+" & _
			"'"&$g_sNameOfSheet_3_GraphMain &"'!"& _Excel_ColumnToLetter($g_iS3EventValueColumn)& $g_iS3EventDistanceRow & "*" & "'"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2EventFirstColumn)&$g_iS2EventSecondHeaderRow+1+$l_iCntr
		$l_sFromTimeFormula = "='"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2EventFromColumn)&$g_iS2EventSecondHeaderRow+1+$l_iCntr
		$l_sToTimeFormula = "='"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2EventToColumn)&$g_iS2EventSecondHeaderRow+1+$l_iCntr
		$l_aEventArray[$l_iCntr*$g_iS5Event_ForLong_NumberOfRowsInSubTable +1][0] = $l_sFromTimeFormula
		$l_aEventArray[$l_iCntr*$g_iS5Event_ForLong_NumberOfRowsInSubTable +1][1] = $l_sLevelFormula

		$l_aEventArray[$l_iCntr*$g_iS5Event_ForLong_NumberOfRowsInSubTable +2][0] = $l_sToTimeFormula
		$l_aEventArray[$l_iCntr*$g_iS5Event_ForLong_NumberOfRowsInSubTable +2][1] = $l_sLevelFormula

		$l_iCntr = $l_iCntr + 1
	WEnd

	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_aEventArray, _Excel_ColumnToLetter($g_iS5Event_ForLong_TimeColumn) & $g_iS5Event_ForLong_FirstHeaderRow)

	;Create activityTable for short graph on sheet 5 for chart
	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Lager tabell for aktivitet i utdragsgraf", "Lager tabell")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Making table for activities in graph selection", "Making table")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1

	If $g_sChosenLanguage = "Norsk" Then
		$l_sTextContainter1 = "Utdragsgraf"
	Else
		$l_sTextContainter1 = "Selection graph"
	EndIf
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sTextContainter1, _Excel_ColumnToLetter($g_iS5Activity_ForShort_HeaderColumn)&"1")
	$l_iCntr = 0
	$l_iActivityTableWidth = 3
	$l_iActivityTableHeight = $g_iNumberOfActivities*($g_iS5Activity_ForShort_NumberOfRowsInSubTable)
	;$l_aActivityArray[$l_iActivityTableHeight][$l_iActivityTableWidth]

	While $l_iCntr < $g_iNumberOfActivities
		$l_aActivityArray[$l_iCntr*$g_iS5Activity_ForShort_NumberOfRowsInSubTable +0][0] = "='"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2ActivityFirstColumn)&$g_iS2ActivitySecondHeaderRow+1+$l_iCntr
		$l_aActivityArray[$l_iCntr*$g_iS5Activity_ForShort_NumberOfRowsInSubTable +0][1] = "='"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2ActivityDescriptionColumn)&$g_iS2ActivitySecondHeaderRow+1+$l_iCntr
		$l_sFromTimeFormula = "='"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2ActivityFromColumn)&$g_iS2ActivitySecondHeaderRow+1+$l_iCntr
		$l_sToTimeFormula = "='"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2ActivityToColumn)&$g_iS2ActivitySecondHeaderRow+1+$l_iCntr
		$l_sLevelFormula = "='"&$g_sNameOfSheet_4_GraphSelection &"'!"& _Excel_ColumnToLetter($g_iS4ActivityValueColumn)& $g_iS4ActivityLevelRow & "-" & _
			"'"&$g_sNameOfSheet_4_GraphSelection &"'!"& _Excel_ColumnToLetter($g_iS4ActivityValueColumn)& $g_iS4ActivityDistanceRow & "+" & _
			"'"&$g_sNameOfSheet_4_GraphSelection &"'!"& _Excel_ColumnToLetter($g_iS4ActivityValueColumn)& $g_iS4ActivityDistanceRow & "*" & "'"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2ActivityFirstColumn)&$g_iS2ActivitySecondHeaderRow+1+$l_iCntr
		$l_aActivityArray[$l_iCntr*$g_iS5Activity_ForShort_NumberOfRowsInSubTable +1][0] = $l_sFromTimeFormula
		$l_aActivityArray[$l_iCntr*$g_iS5Activity_ForShort_NumberOfRowsInSubTable +1][1] = $l_sLevelFormula

		$l_aActivityArray[$l_iCntr*$g_iS5Activity_ForShort_NumberOfRowsInSubTable +2][0] = $l_sToTimeFormula
		$l_aActivityArray[$l_iCntr*$g_iS5Activity_ForShort_NumberOfRowsInSubTable +2][1] = $l_sLevelFormula
		$l_iCntr = $l_iCntr + 1
	WEnd

	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_aActivityArray, _Excel_ColumnToLetter($g_iS5Activity_ForShort_TimeColumn) & $g_iS5Activity_ForShort_FirstHeaderRow)

	;Create eventTable for short graph on sheet 5 for chart
	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Lager tabell for hendelser i utdragsgraf", "Lager tabell")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Making table for events in graph selection", "Making table")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1

	$l_iCntr = 0
	$l_iEventTableWidth = 3
	$l_iEventTableHeight = $g_iNumberOfEvents*($g_iS5Event_ForShort_NumberOfRowsInSubTable)
	;$l_aEventArray[$l_iEventTableHeight][$l_iEventTableWidth]

	While $l_iCntr < $g_iNumberOfEvents
		$l_aEventArray[$l_iCntr*$g_iS5Event_ForShort_NumberOfRowsInSubTable +0][0] = "='"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2EventFirstColumn)&$g_iS2EventSecondHeaderRow+1+$l_iCntr
		$l_aEventArray[$l_iCntr*$g_iS5Event_ForShort_NumberOfRowsInSubTable +0][1] = "='"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2EventDescriptionColumn)&$g_iS2EventSecondHeaderRow+1+$l_iCntr


		$l_sLevelFormula = "='"&$g_sNameOfSheet_4_GraphSelection &"'!"& _Excel_ColumnToLetter($g_iS4EventValueColumn)& $g_iS4EventLevelRow & "-" & _
			"'"&$g_sNameOfSheet_4_GraphSelection &"'!"& _Excel_ColumnToLetter($g_iS4EventValueColumn)& $g_iS4EventDistanceRow & "+" & _
			"'"&$g_sNameOfSheet_4_GraphSelection &"'!"& _Excel_ColumnToLetter($g_iS4EventValueColumn)& $g_iS4EventDistanceRow & "*" & "'"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2EventFirstColumn)&$g_iS2EventSecondHeaderRow+1+$l_iCntr
		$l_sFromTimeFormula = "='"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2EventFromColumn)&$g_iS2EventSecondHeaderRow+1+$l_iCntr
		$l_sToTimeFormula = "='"&$g_sNameOfSheet_2_Observations &"'!"&_Excel_ColumnToLetter($g_iS2EventToColumn)&$g_iS2EventSecondHeaderRow+1+$l_iCntr
		$l_aEventArray[$l_iCntr*$g_iS5Event_ForShort_NumberOfRowsInSubTable +1][0] = $l_sFromTimeFormula
		$l_aEventArray[$l_iCntr*$g_iS5Event_ForShort_NumberOfRowsInSubTable +1][1] = $l_sLevelFormula

		$l_aEventArray[$l_iCntr*$g_iS5Event_ForShort_NumberOfRowsInSubTable +2][0] = $l_sToTimeFormula
		$l_aEventArray[$l_iCntr*$g_iS5Event_ForShort_NumberOfRowsInSubTable +2][1] = $l_sLevelFormula

		$l_iCntr = $l_iCntr + 1
	WEnd

	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_aEventArray, _Excel_ColumnToLetter($g_iS5Event_ForShort_TimeColumn) & $g_iS5Event_ForShort_FirstHeaderRow)

	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Legger ut sammendrag for graf", "Lager ark " & $g_sNameOfSheet_3_GraphMain)
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Creating graph summary", "Making sheet " & $g_sNameOfSheet_3_GraphMain)
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	_Fill_Sheet_3_MainGraph()

	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Formaterer sammendrag for graf", "Lager ark " & $g_sNameOfSheet_3_GraphMain)
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Formatting graph summary", "Making sheet " & $g_sNameOfSheet_3_GraphMain)
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	_Format_Sheet_3_MainGraph()

	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Legger ut sammendrag for graf", "Lager ark " & $g_sNameOfSheet_4_GraphSelection)
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Creating graph summary", "Making sheet " & $g_sNameOfSheet_4_GraphSelection)
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	_Fill_Sheet_4_GraphSelection()

	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Formaterer sammendrag for graf", "Lager ark " & $g_sNameOfSheet_4_GraphSelection)
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Formatting graph summary", "Making sheet " & $g_sNameOfSheet_4_GraphSelection)
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	_Format_Sheet_4_GraphSelection()


	;Creating longGraphTable
	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Fyller kolonner for hovedgraf"& @CR & "For store filer kan dette ta lang tid", "Oppdaterer tabell")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Filling columns for main graph"& @CR & "For large files this can take some time", "Updating table")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1

	Local $l_sLongGraphRange = _Excel_ColumnToLetter($g_iS5Column_TimeForLongGraph) & $g_iS5MainHeaderRow & ":" & _Excel_ColumnToLetter($g_iS5Column_PulseForLongGraph) & $g_iS5LastSecondIntervalRow

	$l_sFormula = "=IF(AND(OFFSET($"& _Excel_ColumnToLetter($g_iS5Column_SecondInterval)&"$"&$g_iS5FirstContentRow&";"
	$l_sFormula = $l_sFormula & "(ROW("& _Excel_ColumnToLetter($g_iS5Column_SecondInterval)&$g_iS5FirstContentRow&")-ROW($"& _Excel_ColumnToLetter($g_iS5Column_SecondInterval)&"$"&$g_iS5FirstContentRow&"))*'"&$g_sNameOfSheet_3_GraphMain&"'!$D$"&$g_iS3GraphTimeResolutionRow&";"
	$l_sFormula = $l_sFormula & "0)>='"&$g_sNameOfSheet_3_GraphMain&"'!$"& _Excel_ColumnToLetter($g_iS3TimeSummaryFromColumn)&"$"&$g_iS3TimeSummaryContentRow&";OFFSET($"& _Excel_ColumnToLetter($g_iS5Column_SecondInterval)&"$"&$g_iS5FirstContentRow&";"
	$l_sFormula = $l_sFormula & "(ROW("& _Excel_ColumnToLetter($g_iS5Column_SecondInterval)&$g_iS5FirstContentRow&")-ROW($"& _Excel_ColumnToLetter($g_iS5Column_SecondInterval)&"$"&$g_iS5FirstContentRow&"))*'"&$g_sNameOfSheet_3_GraphMain&"'!$D$"&$g_iS3GraphTimeResolutionRow&";"
	$l_sFormula = $l_sFormula & "0)<='"&$g_sNameOfSheet_3_GraphMain&"'!$"& _Excel_ColumnToLetter($g_iS3TimeSummaryToColumn)&"$"&$g_iS3TimeSummaryContentRow&");OFFSET($"& _Excel_ColumnToLetter($g_iS5Column_SecondInterval)&"$"&$g_iS5FirstContentRow&";"
	$l_sFormula = $l_sFormula & "(ROW("& _Excel_ColumnToLetter($g_iS5Column_SecondInterval)&$g_iS5FirstContentRow& ")-ROW($"& _Excel_ColumnToLetter($g_iS5Column_SecondInterval)&"$"&$g_iS5FirstContentRow&"))*'"&$g_sNameOfSheet_3_GraphMain&"'!$D$"&$g_iS3GraphTimeResolutionRow&";0); NA())"
	;$l_sFormula = "=$" & _Excel_ColumnToLetter($g_iS5Column_TimeForLongGraph) & "$" & $g_iS5FirstContentRow & "+TIME(0;0;ROW(" & _Excel_ColumnToLetter($g_iS5Column_TimeForLongGraph) & $g_iS5FirstContentRow + 1 & ")-ROW($" & _Excel_ColumnToLetter($g_iS5Column_TimeForLongGraph) & "$" & $g_iS5FirstContentRow & "))"
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_TimeForLongGraph) & $g_iS5FirstContentRow)

	;$l_sFormula = "=VLOOKUP([@["&$g_sS5TimeForLongGraph_ColumnName&"]];"&$g_sS5MainTable_Name&"[[#All];["&$g_sS5CorrectedTimeForMainTable_ColumnName&"]:["&$g_sS5PulseForMainTable_ColumnName&"]];2)"
	$l_sFormula = "=IF(ISBLANK(VLOOKUP([@["&$g_sS5TimeForLongGraph_ColumnName&"]];"&$g_sS5MainTable_Name&"[[#All];["&$g_sS5CorrectedTimeForMainTable_ColumnName&"]:["&$g_sS5PulseForMainTable_ColumnName&"]];2));NA();" & _
		"VLOOKUP([@["&$g_sS5TimeForLongGraph_ColumnName&"]];"&$g_sS5MainTable_Name&"[[#All];["&$g_sS5CorrectedTimeForMainTable_ColumnName&"]:["&$g_sS5PulseForMainTable_ColumnName&"]];2))"
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_PulseForLongGraph) & $g_iS5FirstContentRow)



 	;Creating shortGraphTable
	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Fyller kolonner for kortgraf"& @CR & "For store filer kan dette ta lang tid", "Oppdaterer tabell")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "Filling columns for selection graph"& @CR & "For large files this can take some time", "Updating table")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1

 	Local $l_sShortGraphRange = _Excel_ColumnToLetter($g_iS5Column_TimeForShortGraph) & $g_iS5MainHeaderRow & ":" & _Excel_ColumnToLetter($g_iS5Column_PulseForShortGraph) & $g_iS5LastSecondIntervalRow

 	;$l_sFormula = "=$" & _Excel_ColumnToLetter($g_iS5Column_TimeForShortGraph) & "$" & $g_iS5FirstContentRow & "+TIME(0;0;ROW(" & _Excel_ColumnToLetter($g_iS5Column_TimeForShortGraph) & $g_iS5FirstContentRow + 1 & ")-ROW($" & _Excel_ColumnToLetter($g_iS5Column_TimeForShortGraph) & "$" & $g_iS5FirstContentRow & "))"
	;=IF(AND(OFFSET(INDIRECT($J$5);((ROW(J7)-ROW($J$7))*'Utdrag av graf'!$D$4);0)>='Utdrag av graf'!$D$3;OFFSET(INDIRECT($J$5);((ROW(J7)-ROW($J$7))*'Utdrag av graf'!$D$4);0)<='Utdrag av graf'!$E$3);OFFSET(INDIRECT($J$5);((ROW(J7)-ROW($J$7))*'Utdrag av graf'!$D$4);0); NA())

	Local $l_partFormula = "OFFSET(INDIRECT($"&_Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value) &"$"& $g_iS5Row_ShortGraphFirstValidCell  & ");"& _
		"((ROW("&_Excel_ColumnToLetter($g_iS5Column_TimeForShortGraph) & $g_iS5FirstContentRow&")-ROW($" &_Excel_ColumnToLetter($g_iS5Column_TimeForShortGraph) & "$" & $g_iS5FirstContentRow &"))"& _
		"*'"&$g_sNameOfSheet_4_GraphSelection&"'!$"&_Excel_ColumnToLetter($g_iS4GraphTimeResolutionInputColumn) &"$"& $g_iS4GraphTimeResolutionRow &");0)"
	$l_sFormula = "=IF(AND(" &$l_partFormula&">='"&$g_sNameOfSheet_4_GraphSelection&"'!$"& _Excel_ColumnToLetter($g_iS4GraphTimeSelectionFromColumn) &"$"& $g_iS4GraphTimeSelectionRow& _
		";" & $l_partFormula & "<='"&$g_sNameOfSheet_4_GraphSelection&"'!$"& _Excel_ColumnToLetter($g_iS4GraphTimeSelectionToColumn) &"$"& $g_iS4GraphTimeSelectionRow&");" &$l_partFormula &"; NA())"

 	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_TimeForShortGraph) & $g_iS5FirstContentRow)

;~ 	$g_oExcel.Application.AutoCorrect.AutoFillFormulasInLists = False
;~  	$l_sFormula = "=TIME(" & _Excel_ColumnToLetter($g_iS5Column_CorrectedHour) & $g_iS5FirstContentRow & ";" & _Excel_ColumnToLetter($g_iS5Column_Minute) & $g_iS5FirstContentRow & ";" & _Excel_ColumnToLetter($g_iS5Column_Second) & $g_iS5FirstContentRow & ")"
;~  	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_TimeForShortGraph) & $g_iS5FirstContentRow)
;~  	$g_oExcel.Application.AutoCorrect.AutoFillFormulasInLists = True

	$l_sFormula = "=IF(ISBLANK(VLOOKUP([@["&$g_sS5TimeForShortGraph_ColumnName&"]];"&$g_sS5MainTable_Name&"[[#All];["&$g_sS5CorrectedTimeForMainTable_ColumnName&"]:["&$g_sS5PulseForMainTable_ColumnName&"]];2));NA();" & _
		"VLOOKUP([@["&$g_sS5TimeForShortGraph_ColumnName&"]];"&$g_sS5MainTable_Name&"[[#All];["&$g_sS5CorrectedTimeForMainTable_ColumnName&"]:["&$g_sS5PulseForMainTable_ColumnName&"]];2))"
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_PulseForShortGraph) & $g_iS5FirstContentRow)


	Sleep(1000)
	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "", "Lagrer arbeidsbok")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "", "Saving workbook")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	$l_iResult = _Excel_BookSave($g_oWorkbook)
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSave", "Error saving workbook '" & $l_sXLSX & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	Else
		;MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSave", "Workbook successfully saved as '" & $l_sXLSX & "'.")
	EndIf

	;Insert Average, standard deviation above long graph table
	If $g_sChosenLanguage = "Norsk" Then
		$l_sTextContainter1  = "Sammendrag hovedgraf"
		$l_sTextContainter2  = "Gjennomsnitt"
	Else
		$l_sTextContainter1  = "Summary main graph"
		$l_sTextContainter2  = "Average"
	EndIf
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sTextContainter1, _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header) & $g_iS5Row_LongGraphSummary_Header)
	;Average:   =AVERAGEIF(LongGraphTable[Puls lang graf];"<>#N/A")
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sTextContainter2, _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header) & $g_iS5Row_LongGraphAverage_Header)
	Local $l_sLongPulseRange = "$"& _Excel_ColumnToLetter($g_iS5Column_PulseForLongGraph) & "$"& $g_iS5FirstContentRow &":$" & _Excel_ColumnToLetter($g_iS5Column_PulseForLongGraph) & "$"& $g_iS5LastSecondIntervalRow
	;$l_sFormula =  "=AVERAGEIF("&$g_sS5LongGraphTable_Name &"[" & $g_sS5PulseForLongGraph_ColumnName & "];"& Chr(34) &"<>#N/A"&Chr(34)&")"
	$l_sFormula =  "=AVERAGEIF("&$l_sLongPulseRange  & ";"& Chr(34) &"<>#N/A"&Chr(34)&")"
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Value) & $g_iS5Row_LongGraphAverage_FirstValue &":"& _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Value) & $g_iS5Row_LongGraphAverage_LastValue)
	$l_sFormula = "="& _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5FirstContentRow
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header) & $g_iS5Row_LongGraphAverage_FirstValue)
	$l_sFormula = "="& _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5LastSecondIntervalRow
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header) & $g_iS5Row_LongGraphAverage_LastValue)

 	;STDEV.P: =AGGREGATE( 8;6;LongGraphTable[Puls lang graf])
	If $g_sChosenLanguage = "Norsk" Then
		$l_sTextContainter1  = "Gjennomsnitt pluss ett standardavvik"
		$l_sTextContainter2  = "Gjennomsnitt pluss to standardavvik"
	Else
		$l_sTextContainter1  = "Average plus one st.dev"
		$l_sTextContainter2  = "Average plus two st.dev"
	EndIf
 	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sTextContainter1, _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header) & $g_iS5Row_LongGraphOneStDev_Header)
 ;	_Excel_ColumnToLetter($g_iS5Column_Info_LongGraph_OneStDev) & $g_iS5Row_LongGraph_OneStDev)
 	;$l_sFormula = "=AGGREGATE( 8;6;" & $g_sS5LongGraphTable_Name &"[" & $g_sS5PulseForLongGraph_ColumnName & "]) + " &_Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Value) & $g_iS5Row_LongGraphOneStDev_FirstValue
	$l_sFormula = "=AGGREGATE( 8;6;"&$l_sLongPulseRange  & ") + " &_Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Value) & $g_iS5Row_LongGraphAverage_FirstValue
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Value) & $g_iS5Row_LongGraphOneStDev_FirstValue &":"& _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Value) & $g_iS5Row_LongGraphOneStDev_LastValue)
	$l_sFormula = "="& _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5FirstContentRow
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header) & $g_iS5Row_LongGraphOneStDev_FirstValue)
	$l_sFormula = "="& _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5LastSecondIntervalRow
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header) & $g_iS5Row_LongGraphOneStDev_LastValue)

 	;2*STDEV.P: =AGGREGATE( 8;6;LongGraphTable[Puls lang graf])
 	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sTextContainter2, _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header) & $g_iS5Row_LongGraphTwoStDev_Header)
 ;	_Excel_ColumnToLetter($g_iS5Column_Info_LongGraph_TwoStDev) & $g_iS5Row_LongGraph_TwoStDev)
 	;$l_sFormula = "=AGGREGATE( 8;6;" & $g_sS5LongGraphTable_Name &"[" & $g_sS5PulseForLongGraph_ColumnName & "]) + " &_Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Value) & $g_iS5Row_LongGraphTwoStDev_FirstValue
	$l_sFormula = "=2*AGGREGATE( 8;6;"&$l_sLongPulseRange  & ") + " &_Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Value) & $g_iS5Row_LongGraphAverage_FirstValue
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Value) & $g_iS5Row_LongGraphTwoStDev_FirstValue &":"& _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Value) & $g_iS5Row_LongGraphTwoStDev_LastValue)
	$l_sFormula = "="& _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5FirstContentRow
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header) & $g_iS5Row_LongGraphTwoStDev_FirstValue)
	$l_sFormula = "="& _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5LastSecondIntervalRow
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_LongGraphSummary_Header) & $g_iS5Row_LongGraphTwoStDev_LastValue)

	;Insert Average, standard deviation above Short graph table
	If $g_sChosenLanguage = "Norsk" Then
		$l_sTextContainter1  = "Sammendrag utdragsgraf"
		$l_sTextContainter2  = "Gjennomsnitt"
	Else
		$l_sTextContainter1  = "Summary main graph"
		$l_sTextContainter2  = "Average"
	EndIf
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sTextContainter1, _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header) & $g_iS5Row_ShortGraphSummary_Header)
	;Average:   =AVERAGEIF(ShortGraphTable[Puls lang graf];"<>#N/A")
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sTextContainter2, _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header) & $g_iS5Row_ShortGraphAverage_Header)
	Local $l_sShortPulseRange = "$"& _Excel_ColumnToLetter($g_iS5Column_PulseForShortGraph) & "$"& $g_iS5FirstContentRow &":$" & _Excel_ColumnToLetter($g_iS5Column_PulseForShortGraph) & "$"& $g_iS5LastSecondIntervalRow
	;$l_sFormula =  "=AVERAGEIF("&$g_sS5ShortGraphTable_Name &"[" & $g_sS5PulseForShortGraph_ColumnName & "];"& Chr(34) &"<>#N/A"&Chr(34)&")"
	$l_sFormula =  "=AVERAGEIF("&$l_sShortPulseRange  & ";"& Chr(34) &"<>#N/A"&Chr(34)&")"
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value) & $g_iS5Row_ShortGraphAverage_FirstValue &":"& _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value) & $g_iS5Row_ShortGraphAverage_LastValue)
	$l_sFormula = "="& _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5FirstContentRow
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header) & $g_iS5Row_ShortGraphAverage_FirstValue)
	$l_sFormula = "="& _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5LastSecondIntervalRow
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header) & $g_iS5Row_ShortGraphAverage_LastValue)

 	;STDEV.P: =AGGREGATE( 8;6;ShortGraphTable[Puls lang graf])
	If $g_sChosenLanguage = "Norsk" Then
		$l_sTextContainter1  = "Gjennomsnitt pluss ett standardavvik"
		$l_sTextContainter2  = "Gjennomsnitt pluss to standardavvik"
	Else
		$l_sTextContainter1  = "Average plus one st.dev"
		$l_sTextContainter2  = "Average plus two st.dev"
	EndIf
 	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sTextContainter1, _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header) & $g_iS5Row_ShortGraphOneStDev_Header)
 ;	_Excel_ColumnToLetter($g_iS5Column_Info_ShortGraph_OneStDev) & $g_iS5Row_ShortGraph_OneStDev)
 	;$l_sFormula = "=AGGREGATE( 8;6;" & $g_sS5ShortGraphTable_Name &"[" & $g_sS5PulseForShortGraph_ColumnName & "]) + " &_Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value) & $g_iS5Row_ShortGraphOneStDev_FirstValue
	$l_sFormula = "=AGGREGATE( 8;6;"&$l_sShortPulseRange  & ") + " &_Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value) & $g_iS5Row_ShortGraphAverage_FirstValue
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value) & $g_iS5Row_ShortGraphOneStDev_FirstValue &":"& _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value) & $g_iS5Row_ShortGraphOneStDev_LastValue)
	$l_sFormula = "="& _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5FirstContentRow
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header) & $g_iS5Row_ShortGraphOneStDev_FirstValue)
	$l_sFormula = "="& _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5LastSecondIntervalRow
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header) & $g_iS5Row_ShortGraphOneStDev_LastValue)

 	;2*STDEV.P: =AGGREGATE( 8;6;ShortGraphTable[Puls lang graf])
 	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sTextContainter2, _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header) & $g_iS5Row_ShortGraphTwoStDev_Header)
 ;	_Excel_ColumnToLetter($g_iS5Column_Info_ShortGraph_TwoStDev) & $g_iS5Row_ShortGraph_TwoStDev)
 	;$l_sFormula = "=AGGREGATE( 8;6;" & $g_sS5ShortGraphTable_Name &"[" & $g_sS5PulseForShortGraph_ColumnName & "]) + " &_Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value) & $g_iS5Row_ShortGraphTwoStDev_FirstValue
	$l_sFormula = "=2*AGGREGATE( 8;6;"&$l_sShortPulseRange  & ") + " &_Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value) & $g_iS5Row_ShortGraphAverage_FirstValue
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value) & $g_iS5Row_ShortGraphTwoStDev_FirstValue &":"& _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value) & $g_iS5Row_ShortGraphTwoStDev_LastValue)
	$l_sFormula = "="& _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5FirstContentRow
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header) & $g_iS5Row_ShortGraphTwoStDev_FirstValue)
	$l_sFormula = "="& _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5LastSecondIntervalRow
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header) & $g_iS5Row_ShortGraphTwoStDev_LastValue)

	;Inserting constants for shortgraph calculations
	If $g_sChosenLanguage = "Norsk" Then
		$l_sTextContainter1  = "Første gyldige rad"
		$l_sTextContainter2  = "Første gyldige celle"
	Else
		$l_sTextContainter1  = "First valid row"
		$l_sTextContainter2  = "First valid cell"
	EndIf
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sTextContainter1, _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header) & $g_iS5Row_ShortGraphFirstValidRow)
	;$l_sFormula = "=MATCH('"&$g_sNameOfSheet_4_GraphSelection&"'!D3;I7:I2350)+ ROW(I7)-1"
	$l_sFormula = "=MATCH('"&$g_sNameOfSheet_4_GraphSelection&"'!"& _Excel_ColumnToLetter($g_iS4GraphTimeSelectionFromColumn) & $g_iS4GraphTimeSelectionRow &";"& _
		_Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5FirstContentRow&":"&_Excel_ColumnToLetter($g_iS5Column_SecondInterval) & $g_iS5LastSecondIntervalRow&")+ "&$g_iS5FirstContentRow&"-1"
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value) & $g_iS5Row_ShortGraphFirstValidRow)

	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sTextContainter2, _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Header) & $g_iS5Row_ShortGraphFirstValidCell)
 	$l_sFormula = "=" & Chr(34) & _Excel_ColumnToLetter($g_iS5Column_SecondInterval) & Chr(34) & "&" & _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value)&$g_iS5Row_ShortGraphFirstValidRow 	;Chr(34) = "
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_ShortGraphSummary_Value) & $g_iS5Row_ShortGraphFirstValidCell)


;~ 	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, "Gjennomsnitt pluss to standardavvik", _Excel_ColumnToLetter($g_iS5Column_Info_LongGraph_TwoStDev) & $g_iS5Row_LongGraph_TwoStDev	)
;~ 	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, $l_sFormula, _Excel_ColumnToLetter($g_iS5Column_Input_LongGraph_TwoStDev) & $g_iS5Row_LongGraph_TwoStDev)
;~ 	;Format cells to only  two decimals
;~ 	$g_oWorkbook.Sheets($g_sNameOfSheet_5_CalculationsMainGraph).Range(_Excel_ColumnToLetter($g_iS5Column_Info_LongGraph_Average) & $g_iS5Row_LongGraph_Average & ":" & _Excel_ColumnToLetter($g_iS5Column_Info_LongGraph_Average) & $g_iS5Row_LongGraph_TwoStDev).NumberFormat = "#,##0.00"

	$g_iHourMin = _Excel_RangeRead($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, _Excel_ColumnToLetter($g_iS5Column_CorrectedHour) & $g_iS5FirstContentRow, 1) ;Hour
	$g_iMinuteMin = _Excel_RangeRead($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, _Excel_ColumnToLetter($g_iS5Column_Minute) & $g_iS5FirstContentRow, 1) ;Minute
	$g_iSecondMin = _Excel_RangeRead($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, _Excel_ColumnToLetter($g_iS5Column_Second) & $g_iS5FirstContentRow, 1) ;Second
	$g_iHourMax = _Excel_RangeRead($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, _Excel_ColumnToLetter($g_iS5Column_CorrectedHour) & $g_iS5LastContentRow, 1) ;Hour
	$g_iMinuteMax = _Excel_RangeRead($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, _Excel_ColumnToLetter($g_iS5Column_Minute) & $g_iS5LastContentRow, 1) ;Minute
	$g_iSecondMax = _Excel_RangeRead($g_oWorkbook, $g_sNameOfSheet_5_CalculationsMainGraph, _Excel_ColumnToLetter($g_iS5Column_Second) & $g_iS5LastContentRow, 1) ;Second

	;Inserting suggestd time range for graph selection
	Local $l_sFirstTimeString = $g_iHourMin & ":" & $g_iMinuteMin & ":" & $g_iSecondMin
	$g_iS4SelectedTo_Hour = $g_iHourMin
	$g_iS4SelectedTo_Minute  = $g_iMinuteMin + Round($g_iS4MaxLengthOfShortGraph/60,0)
	$g_iS4SelectedTo_Second = $g_iSecondMin

	IF $g_iS4SelectedTo_Minute > 59 Then
		$g_iS4SelectedTo_Minute = $g_iS4SelectedTo_Minute - 60
		$g_iS4SelectedTo_Hour = $g_iHourMin + 1
	EndIf

	Local $l_sLastTimeString = $g_iS4SelectedTo_Hour & ":" & $g_iS4SelectedTo_Minute & ":" & $g_iS4SelectedTo_Second
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_4_GraphSelection, $l_sFirstTimeString, _Excel_ColumnToLetter($g_iS4GraphTimeSelectionFromColumn) & $g_iS4GraphTimeSelectionRow)
	_Excel_RangeWrite($g_oWorkbook, $g_sNameOfSheet_4_GraphSelection, $l_sLastTimeString, _Excel_ColumnToLetter($g_iS4GraphTimeSelectionToColumn) & $g_iS4GraphTimeSelectionRow)

	If $g_sChosenLanguage = "Norsk" Then
		$l_sTextContainter1  = "Legger ut graf"
		$l_sTextContainter2  = "Lager ark "
	Else
		$l_sTextContainter1  = "Creating graph"
		$l_sTextContainter2  = "Making sheet "
	EndIf

	ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), $l_sTextContainter1, $l_sTextContainter2 & $g_sNameOfSheet_3_GraphMain)
	$g_iProgressCounter = $g_iProgressCounter + 1
	_Add_MainGraph_To_Sheet_2()

	ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), $l_sTextContainter1, $l_sTextContainter2 & $g_sNameOfSheet_4_GraphSelection)
	$g_iProgressCounter = $g_iProgressCounter + 1
	_Add_SelectionGraph_To_Sheet_3()

	$g_oWorkbook.Sheets($g_sNameOfSheet_5_CalculationsMainGraph).Columns($g_iS5Column_TimeForShortGraph).NumberFormat = "[$-x-systime]h:mm:ss AM/PM"
	$g_oWorkbook.Sheets($g_sNameOfSheet_5_CalculationsMainGraph).Columns($g_iS5Column_TimeForLongGraph).NumberFormat = "[$-x-systime]h:mm:ss AM/PM"
	$g_oWorkbook.Sheets($g_sNameOfSheet_5_CalculationsMainGraph).Columns($g_iS5Column_SecondInterval).NumberFormat = "[$-x-systime]h:mm:ss AM/PM"

	Sleep(1000)
	If $g_sChosenLanguage = "Norsk" Then
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "", "Lagrer arbeidsbok")
	Else
		ProgressSet(100 * (Round(($g_iProgressCounter + 1) / $g_iProgressMaxCount, 2)), "", "Saving workbook")
	EndIf
	$g_iProgressCounter = $g_iProgressCounter + 1
	$l_iResult = _Excel_BookSave($g_oWorkbook)
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSave", "Error saving workbook '" & $l_sXLSX & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	Else
		;MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSave", "Workbook successfully saved as '" & $l_sXLSX & "'.")
	EndIf

	ProgressSet (100*(Round(($g_iProgressCounter +1) /$g_iProgressMaxCount,2)) , "" , "Ferdig")


	;MsgBox(0, "Progress", $g_iProgressCounter)
	ProgressOff()
	$g_iProgressCounter = 0
	If $g_sChosenLanguage = "Norsk" Then
		$l_sTextContainter1  = "Ferdig"
		$l_sTextContainter2  = "Omforming er ferdig. Du kan velge å omgjøre en ny fil eller du kan avslutte programmet." &@CRLF&@CRLF &"Hvis du velger å omgjøre en ny fil bør du lukke alle åpne Excel-filer før du starter"
	Else
		$l_sTextContainter1  = "Finished"
		$l_sTextContainter2  = "Transforming completed. You can choose to transform a new file or close the program." &@CRLF&@CRLF &"If you choose to transform a new file you should close all open Excel files before beginning transformation"
	EndIf
	MsgBox($MB_SYSTEMMODAL, $l_sTextContainter1, $l_sTextContainter2)
	Return 1
EndFunc   ;==>OnCreateExcel

Func MyErrFunc($l_oErrorHandler)
    ; Important: the error object variable MUST be named $l_oErrorHandler
    $ErrorScriptline = $l_oErrorHandler.scriptline
    $ErrorNumber = $l_oErrorHandler.number
    $ErrorNumberHex = Hex($l_oErrorHandler.number, 8)
    $ErrorDescription = StringStripWS($l_oErrorHandler.description, 2)
    $ErrorWinDescription = StringStripWS($l_oErrorHandler.WinDescription, 2)
    $ErrorSource = $l_oErrorHandler.Source
    $ErrorHelpFile = $l_oErrorHandler.HelpFile
    $ErrorHelpContext = $l_oErrorHandler.HelpContext
    $ErrorLastDllError = $l_oErrorHandler.LastDllError
    $ErrorOutput = ""
    $ErrorOutput &= "--> COM Error Encountered in " & @ScriptName & @CR
    $ErrorOutput &= "----> $ErrorScriptline = " & $ErrorScriptline & @CR
    $ErrorOutput &= "----> $ErrorNumberHex = " & $ErrorNumberHex & @CR
    $ErrorOutput &= "----> $ErrorNumber = " & $ErrorNumber & @CR
    $ErrorOutput &= "----> $ErrorWinDescription = " & $ErrorWinDescription & @CR
    $ErrorOutput &= "----> $ErrorDescription = " & $ErrorDescription & @CR
    $ErrorOutput &= "----> $ErrorSource = " & $ErrorSource & @CR
    $ErrorOutput &= "----> $ErrorHelpFile = " & $ErrorHelpFile & @CR
    $ErrorOutput &= "----> $ErrorHelpContext = " & $ErrorHelpContext & @CR
    $ErrorOutput &= "----> $ErrorLastDllError = " & $ErrorLastDllError
    MsgBox(0,"COM Error", $ErrorOutput)
    SetError(1)
    Return
EndFunc  ;==>MyErrFunc
Func _ReadConfigurationFile()
	;IniWrite(@ScriptDir&"\"&$g_sIniFilename, "Tittel", "Konfigurering av Forstå-meg filomformer")
	;MsgBox(0, "INI-filen heter:", @ScriptDir&"\"&$g_sIniFilename)

	$g_sIniValue_RawFolder = IniRead(@ScriptDir & "\" & $g_sIniFilename, $g_sIniSection, $g_sIniKey_Raw, "Default")
	$g_sIniValue_XMLFolder = IniRead(@ScriptDir & "\" & $g_sIniFilename, $g_sIniSection, $g_sIniKey_XML, "Default")
	$g_sIniValue_ExcelFolder = IniRead(@ScriptDir & "\" & $g_sIniFilename, $g_sIniSection, $g_sIniKey_Excel, "Default")
	$g_sIniValue_GPSBabelFolder = IniRead(@ScriptDir & "\" & $g_sIniFilename, $g_sIniSection, $g_sIniKey_GPSBabel, "Default")

EndFunc   ;==>_ReadConfigurationFile

Func _TransformFile($l_oInputFile, $l_oOutputFile)
	Local $l_sCMD = "C:\Program Files (x86)\GPSBabel\gpsbabel -t -i garmin_fit,allpoints=1 -f "
	Local $l_sOutputParameters = " -o gtrnctr -F "
	Local $l_sRunString = $l_sCMD & '"' & $l_oInputFile & '"' & $l_sOutputParameters & '"' & $l_oOutputFile & '"'
	RunWait($l_sRunString, "", @SW_SHOWDEFAULT)
	If FileExists($l_oOutputFile) = 0 Then
		If $g_sChosenLanguage = "Norsk" Then
			MsgBox($MB_SYSTEMMODAL, "Feil", "Noe gikk galt under opprettelse av filen." & @CRLF & $l_oOutputFile & "Vennligst kontroller at valgt filnavn og bane er korrekt.")
		Else
			MsgBox($MB_SYSTEMMODAL, "Error", "Something went wrong when creating the file." & @CRLF & $l_oOutputFile & "Please check that the given file name and path are correct.")
		EndIf
	Else
		;MsgBox($MB_SYSTEMMODAL, "Vellykket omforming", "Filen "& @CRLF&$l_oFile& @CRLF&" er nå opprettet")
	EndIf
EndFunc   ;==>_TransformFile

Func OnExit()
	Exit
EndFunc   ;==>OnExit
