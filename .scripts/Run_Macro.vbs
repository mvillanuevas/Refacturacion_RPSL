On Error Resume Next

' Argumentos: libro origen, libro destino
Set objArgs = WScript.Arguments

WorkBookPath = objArgs(0)
WorkBookName = objArgs(1)
MacroName = objArgs(2)

' WorkBookPath = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\Timbrado\Refacturacion_regular_v2.xlsm"
' WorkBookName = "Refacturacion_regular_v2.xlsm"
' MacroName = "paso1_Formulas"
 
Set objExcel = CreateObject("Excel.Application")

'Parámetro para indicar si se quiere visible la aplicación de Excel
objExcel.Application.Visible = True
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = True
'Parámetro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False
 
Set objWorkBookPath = objExcel.Workbooks.Open(WorkBookPath, 0)
 
' Ejecutar la macro
objExcel.Run WorkBookName & "!" & MacroName
 
' Guardar cambios y cerrar
objWorkBookPath.Save
objWorkBookPath.Close False
objExcel.Quit
 
 
If Err.Number <> 0 Then
    WScript.Echo "Error: " & Err.Description
Else
    WScript.Echo "Macro ejecutada correctamente."
End If