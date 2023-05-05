'##################################################################################################################################
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
' Run  UFT  in minimize mode
Set QtApp = CreateObject("QuickTest.Application") 
QtApp.WindowState = "Minimized"

'Variable Declaration
Dim varStatusMessage, varStatusText, varMaterial

'Lauch SAP Transaction code
SAPGuiSession("Session").Reset "VA01"

'Enter Data into the fields
SAPGuiSession("Session").SAPGuiWindow("Create Sales Documents").SAPGuiEdit("*Order Type").Set Parameter("ipOrderType")'"OR"
SAPGuiSession("Session").SAPGuiWindow("Create Sales Documents").SAPGuiEdit("Sales Organization").Set Parameter("ipSalesOrganisation")'"1710"
SAPGuiSession("Session").SAPGuiWindow("Create Sales Documents").SAPGuiEdit("Distribution Channel").Set Parameter("ipDistributionChannel")'"10"
SAPGuiSession("Session").SAPGuiWindow("Create Sales Documents").SAPGuiEdit("Division").Set Parameter("ipDivision")'"00"
'Click on Enter button
SAPGuiSession("Session").SAPGuiWindow("Create Sales Documents").SAPGuiButton("Continue").Click
'Enter required details
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiEdit("Cust. Reference").Set Parameter("ipCustReference")'"450000019998"
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiEdit("Cust. Ref. Date").Set Date'Parameter("ipCustReferenceDate")'"04/05/2023"
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiEdit("Sold-To Party").Set Parameter("ipSoldToParty")'"EWM17-CU02"
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiEdit("Ship-To Party").Set Parameter("ipShipToParty")'"EWM17-CU02"
'Click on Enter button
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SendKey ENTER
'Enter data in the all items table
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiTable("All Items").SetCellData 1,"Item",Parameter("ipItem")'"10"
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiTable("All Items").SetCellData 1,"Material",Parameter("ipMaterial")'"EWMS4-01"
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiTable("All Items").SetCellData 1,"Order Quantity",Parameter("ipQty")'"1"
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiTable("All Items").SetCellData 1,"Un",Parameter("IpUnit")'"PC"
'Click on Save button
SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiButton("Save").Click
'Inlcuded to manage the pop-up window
If SAPGuiSession("Session").SAPGuiWindow("Open quotations for item").SAPGuiButton("Continue").Exist (5) Then
	SAPGuiSession("Session").SAPGuiWindow("Open quotations for item").SAPGuiButton("Continue").Click
End If
'inlcuded to manage material availablity window
 If SAPGuiSession("Session").SAPGuiWindow("Standard Order: Availability").SAPGuiButton("Continue").Exist (5) Then
 	SAPGuiSession("Session").SAPGuiWindow("Standard Order: Availability").SAPGuiButton("Continue").Click
 End If

 'Validate and retrieve sales order from statusbar
 If SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiStatusBar("StatusBar").Exist Then
 	varStatusMessage = SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiStatusBar("StatusBar").GetROProperty("messagetype")
 	If varStatusMessage = "S" Then
 		varStatusText = SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiStatusBar("StatusBar").GetROProperty("text")
 		Parameter ("opSalesOrderNumber") = SAPGuiSession("Session").SAPGuiWindow("Create Standard Order:").SAPGuiStatusBar("StatusBar").GetROProperty("item2")
 		Reporter.ReportEvent micPass, "Create Sales Order", "Sales Order created with the document number : "& Parameter ("opSalesOrderNumber")
 	  Else
 	  	Reporter.ReportEvent micFail, "Create Sales Order", "Sales order creation was failed, please check your entries"
 	End If
   Else
   	Reporter.ReportEvent micFail, "Create Sales Order", "Sales order creation failed, no status message displayed in the statusbar"
 End If

 'Assign Materrial to a output parameter
 varMaterial = Parameter("ipMaterial")
 Parameter ("opMaterial") = varMaterial

 '***********************************************End of Script*******************************************************
