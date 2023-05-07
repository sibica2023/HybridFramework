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

'Give the path of the Data file
Environment.Value("strFilePath") =  "C:\Users\demo\Documents\UFT One\HybridFramework\DataSheet\OrderToCash.xlsx" 

'Create an Excel Object and open the input data file
 Set xlObj = CreateObject("Excel.Application") 
 xlObj.WorkBooks.Open Environment.Value("strFilePath") 
 xlObj.DisplayAlerts = True
 xlObj.Visible = True
 Set xlWB = xlObj.ActiveWorkbook 
 Set xlSheet = xlWB.WorkSheets("SAPLogOn") 
 
'Set current row
intCurrentRow = Parameter("Row")

If Ucase (GetColValue("ExecuteIteration"))="TRUE" Then
	'Variable Declaration
	Dim varServerDescription : varServerDescription = GetColValue("ipServerDescription")
	Dim intClient : intClient = GetColValue("ipClient")
	Dim varUsername : varUsername = GetColValue("ipUsername")
	Dim varPassword : varPassword = GetColValue("ipPassword")
	Dim varLanguage : varLanguage = GetColValue("ipLanguage")
	
	'SAP Logon
	SAPGuiUtil.AutoLogon varServerDescription, intClient, varUsername, varPassword, varLanguage
	Wait (5)
	
	'Validation
	If SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access").Exist (5) Then
		Reporter.ReportEvent micPass, "SAP Logon", "SAP logon is sucessful"
		Parameter ("bIterationStatus") = "PASS"
	    Else
	    	Reporter.ReportEvent micFail, "SAP Logon", "SAP Logon failed, please check your entries"
	 End  If
End  If
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf1.xml_;_
 '***********************************************End of Script*******************************************************

 'Function Name  GetColValue
         'Description  : Returns column no. based on column name

		 Public Function GetColValue(stringCN)
			intColumnCnt=xlSheet.usedrange.Entirecolumn.count
            For i = 1 to intColumnCnt
				If (stringCN = xlSheet.Cells(1,i).value) Then
					If   xlSheet.Cells(intCurrentRow,i).value <> "" Then
						GetColValue = xlSheet.Cells(intCurrentRow,i).value
					Else
						Reporter.ReportEvent micFail,"Input Data Validation", stringCN & " Value in datasheet  is empty " 
					End If					
                    Exit for
				End If
			Next
		 End Function
'--------------------------------------------------------------------------------------------------------------------------

'===================================================================================
' Function Name: SetXLVal
' Description  : To set Value to XL sheet
' Return Value : Column name, Row no and cell value
Function SetXLVal(ColumnName,RowNo,CellValue)
 intColumnCnt=xlSheet.usedrange.Entirecolumn.count
 For i = 1 to intColumnCnt
  If (ColumnName = cstr(xlSheet.Cells(1,i).value)) Then
   ColValue = i
   Exit for
  End If
 Next
 xlSheet.Cells(RowNo,ColValue)=CellValue
end Function

