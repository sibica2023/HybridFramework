'##################################################################################################################################
'Project Name : JCI Project
'File Name: PME_GL_Post FI document
'Description:Copy Actual to Plan
'Developed  by/Date: Sibi/ 25.02.2012
'Version No:0.1
'Data File Name: Excel Sheet
'Mandatory Fields:
'Input Parameters Used:  From data sheets
'Output Parameters Used: bIterationRunStatus
'Reviewed by/Review Date: 

'*******************************************************************************Modification history***********************************************************************************
'S.No___________________________Modified by__________________________Modified Date__________________________Reason____________________

'****************************************************************************************************************************************************************************************' 

'####################################################################################################################################

' To close all the excels sheets present in the system
SystemUtil.CloseProcessByName("Excel.exe")

' Run  QTP  in minimize mode
Set QtApp = CreateObject("QuickTest.Application") 
QtApp.WindowState = "Minimized"

'Give the path of the UserDefinedFunctions.vbs file and execute
  strVbsPath = "C:\Automation\Lib Fun\Library Function.qfl" 
  ExecuteFile strVbsPath

'Give the path of the Data file
Environment.Value("strFilePath") =  "C:\Data Sheet\PME_GL_Post FI document.xlsx" 

'Included to generate the screenshots
Environment.Value("reportPath") = hour(now)&minute(now)&second(now)

'Create an Excel Object and open the input data file
Set xlObj = CreateObject("Excel.Application") 
 xlObj.WorkBooks.Open Environment.Value("strFilePath") 
 Set xlWB = xlObj.ActiveWorkbook 
 Set xlSheet = xlWB.WorkSheets("SAPLogin") 

Environment.Value("AllRows") = xlSheet.UsedRange.Rows.Count
xlWB.Save
xlObj.Quit

For intcurrentRow=2 to Environment.Value("AllRows")

		RunAction "Action1 [SAP Login]", oneIteration,intcurrentRow,RunStatusLogin     'Login

			If  RunStatusLogin = "Pass" Then
				 RunAction "Action1 [FB01_Post_document]", oneIteration,intcurrentRow,RunStatusPostDocu    'Post FI Document
			End If

			If  RunStatusPostDocu = "Pass" Then
  				 RunAction "Action1 [FB03_Display Document Overview]", oneIteration,intcurrentRow,0,RunStatusDisplay   'Display Document
			End If

		RunAction "Action1 [SAP Logoff]", oneIteration

Next


'*******************************************************************************End of Script******************************************************************************************

