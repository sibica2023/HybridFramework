'##################################################################################################################################
'Project Name : 
'File Name: SAP Logoff
'Transaction covered: SAP Easy Access
'Description: This script is used logoff from SAP
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
'Close all SAP active sessions
SAPGuiUtil.CloseConnections

'Validation
If not SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access").Exist (5) Then
 	Reporter.ReportEvent micPass, "SAP Logoff", "SAP logoff is successful"
   Else
   	Reporter.ReportEvent micFail, "SAP Logoff", "SAP logoff is not successful, session is still active"	
End If
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf1.xml_;_
'**********************************************************End of Script***************************************************
