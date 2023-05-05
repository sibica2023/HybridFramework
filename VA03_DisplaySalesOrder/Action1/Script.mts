'##################################################################################################################################
'Project Name : 
'File Name: VA03 Display Sales Order
'Transaction covered: VA03
'Description: This script is used to display sales order
'Developed  by/Date: Sibi C A / 05-Apr-2023
'Version No: 0.1
'Data File Name: NA
'Mandatory Fields:
'Input Parameters Used:  NA
'Output Parameters Used: NA
'Reviewed by/Review Date: 

'**********************************************Modification history********************************************************************
'S.No___________________________Modified by__________________________Modified Date__________________________Reason____________________

'***********************************************************************************************************************************' 

'####################################################################################################################################

' Run  QTP  in minimize mode
SystemUtil.CloseProcessByName("Excel.exe")
Set QtApp = CreateObject("QuickTest.Application") 
QtApp.WindowState = "Minimized"

'Give the path of the UserDefinedFunctions.vbs file and execute
 strVbsPath = "C:\Automation\Lib Fun\Library Function.qfl" 
 ExecuteFile strVbsPath

'Give the path of the Data file
'Environment.Value("strFilePath") =  "C:\Data Sheet\PME_Post vendor invoice with WH tax code ZX.xlsx" 

'intCurrentRow = 2
'intDataSetCnt = 0

'Create an Excel Object and open the input data file
 Set xlObj = CreateObject("Excel.Application") 
 xlObj.WorkBooks.Open Environment.Value("strFilePath") 
 xlObj.DisplayAlerts = False
 Set xlWB = xlObj.ActiveWorkbook 
 Set xlSheet = xlWB.WorkSheets("VA03") 

intCurrentRow = Parameter("Row")

If Ucase (GetColValue("ExecuteIteration"))="TRUE" Then

	'Variable declaration
	Dim intDocumentNo
	Dim intCompanyCode	
	'Enter corresponding transaction code
	SAPGuiSession("Session").Reset "VA03"
	

End  If


'**********************************************************End of Script***************************************************
