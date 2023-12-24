Attribute VB_Name = "aaa_15_populateTlb"
Sub macro_15_populateTlb()
Attribute macro_15_populateTlb.VB_ProcData.VB_Invoke_Func = " \n14"

'PUT THE SHEET NAME DESTINATION
Const shtnam_1 As String = "tbl" '< = Pon tu hoja de destinacíon aqui
Dim addrr_1 As String

Dim ar1 As Range
Dim wks1 As Worksheet

'COPY THE SELECTED CELL
 Set wks1 = ActiveSheet
 addrr_1 = ActiveCell.Address

 'EXIT MACRO IF POSITONED ON DESTINATION SHEET
 If wks1.Name = shtnam_1 Then Exit Sub
 
'GOTO TO TBL SHEET
    Sheets(shtnam_1).Activate
    
'INSERT BLANK ROW
    Rows("2:2").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
 'APPEND FIELDS
    ActiveCell.Offset(0, 1) = Date
    ActiveCell.Offset(0, 2).FormulaR1C1 = "=RANDBETWEEN(0,6)^2" ' < = Reemplaca esta fila con tu formula VLOOKUP

'RETURN TO SELECTION
    wks1.Activate
    Range(addrr_1).Select
    Selection.Copy
    
'GOTO TO TBL SHEET
    Sheets(shtnam_1).Activate
    
'COPY SELECTION TO UNOCUPIED ROW
    Cells(2, 1).Select
    ActiveSheet.Paste
    
End Sub
