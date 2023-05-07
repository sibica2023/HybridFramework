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

'Give the path of the Data file
'Environment.Value("strFilePath") =  "C:\Users\demo\Documents\UFT One\HybridFramework\DataSheet\OrderToCash.xlsx" 

'Create an Excel Object and open the input data file
 Set xlObj = CreateObject("Excel.Application") 
 xlObj.WorkBooks.Open Environment.Value("strFilePath") 
 xlObj.DisplayAlerts = False
 xlObj.Visible = False
 Set xlWB = xlObj.ActiveWorkbook 
 Set xlSheet = xlWB.WorkSheets("VA03") 
 
'Set current row
intCurrentRow = Parameter("Row")

If Ucase (GetColValue("ExecuteIteration"))="TRUE" Then

	'Variable Declaration
	Dim varStatusMessage, varStatusText, varMaterial,opSalesOrderNumber
	'Lauch SAP Transaction code
	SAPGuiSession("Session").Reset "VA03"
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf1.xml_;_
	 'Enter Sales Order number
	SAPGuiSession("Session").SAPGuiWindow("Display Sales Documents").SAPGuiEdit("Order").Set GetColValue("ipSalesOrderNumber") @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf1.xml_;_
	'Click on Search button
	SAPGuiSession("Session").SAPGuiWindow("Display Sales Documents").SAPGuiButton("Search").Click @@ hightlight id_;_2_;_script infofile_;_ZIP::ssf1.xml_;_
	
	'Verify displayed sales order number
	varSalesOrderActual = SAPGuiSession("Session").SAPGuiWindow("Display Standard Order").SAPGuiEdit("Standard Order").GetROProperty ("value") @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf2.xml_;_
	If varSalesOrderActual = GetColValue("ipSalesOrderNumber") Then
		Reporter.ReportEvent micPass, "Display Sales Order", "Sales order displayed and verified successfully"
		Parameter ("bIterationStatus") = "PASS"
	    Else
		 Reporter.ReportEvent micFail, "Display Sales Order", "Display sales order was failed, please check your entries"
	End If
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf2.xml_;_
	 'Navigate back to SAP Easy Access screen
	SAPGuiSession("Session").SAPGuiWindow("Display Standard Order").SAPGuiButton("Back   (F3)").Click @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf3.xml_;_
	SAPGuiSession("Session").SAPGuiWindow("Display Sales Documents").SAPGuiButton("Back   (F3)").Click @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf4.xml_;_
	
End  If

'**********************************************************End of Script***************************************************

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

