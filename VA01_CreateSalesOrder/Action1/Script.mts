﻿'##################################################################################################################################
'Project Name : 
'File Name: VA01-Create Sales Order
'Transaction covered: VA01
'Description: This script is used create sales order in VA01
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
 Set xlSheet = xlWB.WorkSheets("VA01") 
 
'Set current row
intCurrentRow = Parameter("Row")

If Ucase (GetColValue("ExecuteIteration"))="TRUE" Then

	'Variable Declaration
	Dim varStatusMessage, varStatusText, varMaterial,opSalesOrderNumber
	'Lauch SAP Transaction code
	SAPGuiSession("Session").Reset "VA01"

	'Enter Data into the fields @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf1.xml_;_
	SAPGuiSession("Session").SAPGuiWindow("Create Sales Documents").SAPGuiEdit("Order Type").Set GetColValue("ipOrderType")'"OR" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf1.xml_;_
	SAPGuiSession("Session").SAPGuiWindow("Create Sales Documents").SAPGuiEdit("Sales Organization").Set GetColValue("ipSalesOrganisation")'"1710" @@ hightlight id_;_2_;_script infofile_;_ZIP::ssf1.xml_;_
	SAPGuiSession("Session").SAPGuiWindow("Create Sales Documents").SAPGuiEdit("Distribution Channel").Set GetColValue("ipDistributionChannel")'"10" @@ hightlight id_;_3_;_script infofile_;_ZIP::ssf1.xml_;_
	SAPGuiSession("Session").SAPGuiWindow("Create Sales Documents").SAPGuiEdit("Division").Set GetColValue("ipDivision")'"00"
	'Click on Enter button
	SAPGuiSession("Session").SAPGuiWindow("Create Sales Documents").SAPGuiButton("Continue   (Enter)").Click
	'Enter required details @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf2.xml_;_
	SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiEdit("Sold-To Party").Set GetColValue("ipSoldToParty")'"EWM17-CU02" @@ hightlight id_;_2_;_script infofile_;_ZIP::ssf2.xml_;_
	SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiEdit("Ship-To Party").Set  GetColValue("ipShipToParty")'"EWM17-CU02" @@ hightlight id_;_3_;_script infofile_;_ZIP::ssf2.xml_;_
	SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiEdit("Cust. Reference").Set GetColValue("ipCustReference")'"450000019998" @@ hightlight id_;_4_;_script infofile_;_ZIP::ssf2.xml_;_
	SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiEdit("Cust. Ref. Date").Set Date @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf3.xml_;_
	'Click on Enter button
	SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiButton("Enter").Click @@ hightlight id_;_2_;_script infofile_;_ZIP::ssf3.xml_;_
	'Enter data in the all items table
	SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiTable("All Items").SetCellData 1,"Item",GetColValue("ipItem")'"10" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf4.xml_;_
	SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiTable("All Items").SetCellData 1,"Material",GetColValue("ipMaterial")'"EWMS4-01" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf4.xml_;_
	SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiTable("All Items").SetCellData 1,"Order Quantity",GetColValue("ipQty")'"1" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf4.xml_;_
	SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiTable("All Items").SetCellData 1,"Un",GetColValue("IpUnit")'"PC" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf4.xml_;_
	'Click on Enter button @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf4.xml_;_
	SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiButton("Enter").Click
	'Inlcuded to handle information pop-up
	If SAPGuiSession("Session").SAPGuiWindow("Open quotations for item").Exist Then
		SAPGuiSession("Session").SAPGuiWindow("Open quotations for item").SAPGuiButton("Continue").Click
	End If
	If SAPGuiSession("Session").SAPGuiWindow("Standard Order: Availability").Exist Then
		SAPGuiSession("Session").SAPGuiWindow("Standard Order: Availability").SAPGuiButton("Continue").Click
	End If @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf6.xml_;_
	 'Click onSave button
	SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiButton("Save   (Ctrl+S)").Click
	 'Validate and retrieve sales order from statusbar
	 If SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiStatusBar("StatusBar").Exist Then
	 	varStatusMessage = SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiStatusBar("StatusBar").GetROProperty("messagetype")
	 	If varStatusMessage = "S" Then
	 		varStatusText = SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiStatusBar("StatusBar").GetROProperty("text")
	 		opSalesOrderNumber = SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiStatusBar("StatusBar").GetROProperty("item2")
	 		Reporter.ReportEvent micPass, "Create Sales Order", "Sales Order created with the document number : "& opSalesOrderNumber
	 		Parameter("bIterationStatus") = "PASS"
	 	  Else
	 	  	Reporter.ReportEvent micFail, "Create Sales Order", "Sales order creation was failed, please check your entries"
	 	End If
	   Else
	   	Reporter.ReportEvent micFail, "Create Sales Order", "Sales order creation failed, no status message displayed in the statusbar"
	 End If
	
End  If

'These codes are included to save the data into the respective sheets
Set xlSheet = nothing
For Iter = 1 To xlWB.Worksheets.Count
	 If xlWB.Worksheets(Iter).Name = "VA03" Then 
		 Set xlSheet = xlWB.Worksheets(Iter)
		 setxlval "ipSalesOrderNumber",intCurrentRow, opSalesOrderNumber
		 Exit For 
     End If 
Next 
Set xlSheet = nothing
For Iter = 1 To xlWB.Worksheets.Count
	 If xlWB.Worksheets(Iter).Name = "VA01" Then 
		 Set xlSheet = xlWB.Worksheets(Iter)
		 setxlval "opSalesOrderNumber",intCurrentRow, opSalesOrderNumber
		 Exit For 
     End If 
Next 
	 
xlWB.Save
xlObj.Quit
Set xlSheet = nothing

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

