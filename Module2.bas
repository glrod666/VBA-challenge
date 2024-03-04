Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ChDir "C:\Users\gezeu\OneDrive\Desktop"
    ActiveWorkbook.SaveAs Filename:= _
        "https://d.docs.live.net/918cafea4bf26718/Desktop/Multiple_year_stock_data.xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
End Sub
