'##################################################################################################################################
'Project Name : 
'File Name: SAP Logon
'Transaction covered: SAP Easy Access
'Description: This script is used loginto SAP
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
'Close all existing SAP connections
SAPGuiUtil.CloseConnections

' Run  UFT  in minimize mode
Set QtApp = CreateObject("QuickTest.Application") 
QtApp.WindowState = "Minimized"

'Variable Declaration
Dim varServerDescription : varServerDescription = Parameter("ipServerDescription")
Dim intClient : intClient = Parameter("ipClient")
Dim varUsername : varUsername = Parameter("ipUsername")
Dim varPassword : varPassword = Parameter("ipPassword")
Dim varLanguage : varLanguage = Parameter("ipLanguage")

'SAP Logon
SAPGuiUtil.AutoLogon varServerDescription, intClient, varUsername, varPassword, varLanguage
Wait (5)

'Validation
If SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access").Exist (5) Then
	Reporter.ReportEvent micPass, "SAP Logon", "SAP logon is sucessful"
    Else
    	Reporter.ReportEvent micFail, "SAP Logon", "SAP Logon failed, please check your entries"
 End  If
 
 '***********************************************End of Script*******************************************************
