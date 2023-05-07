'##################################################################################################################################
'Project Name : Order To Cash Process Flow
'File Name: Create and display sales order
'Description:This end to end scenario is used to create and display sales order
'Developed  by/Date: Sibi C A/ 06-Apr-2023
'Version No:0.1
'Data File Name: Excel Sheet
'Mandatory Fields:
'Input Parameters Used:  From data sheets
'Output Parameters Used: 
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

'Execute Library Function file


'Give the path of the Data file
Environment.Value("strFilePath") =  "C:\Users\demo\Documents\UFT One\HybridFramework\DataSheet\OrderToCash.xlsx" 

'Included to generate the screenshots
Environment.Value("reportPath") = hour(now)&minute(now)&second(now)

'Create an Excel Object and open the input data file
Set xlObj = CreateObject("Excel.Application") 
 xlObj.WorkBooks.Open Environment.Value("strFilePath") 
 Set xlWB = xlObj.ActiveWorkbook 
 Set xlSheet = xlWB.WorkSheets("SAPLogOn") 

Environment.Value("AllRows") = xlSheet.UsedRange.Rows.Count
xlWB.Save
xlObj.Quit

For intcurrentRow = 2 to Environment.Value("AllRows")

		RunAction "Action1 [SAPLogon]", oneIteration, ,intcurrentRow,RunStatusLogin
			If  RunStatusLogin = "PASS" Then		
				RunAction "Action1 [VA01_CreateSalesOrder]", oneIteration, intcurrentRow,RunStatusCreateSO
			End If
			If  RunStatusCreateSO = "PASS" Then				
				RunAction "Action1 [VA03_DisplaySalesOrder]", oneIteration, intcurrentRow,RunStatusDisplay
			End If
		RunAction "Action1 [SAPLogOff]", oneIteration

Next

'*******************************************************************************End of Script******************************************************************************************

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

