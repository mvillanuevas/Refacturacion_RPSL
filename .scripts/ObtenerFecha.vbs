On Error Resume Next

Set objArgs = WScript.Arguments

excelPath = objArgs(0)


' === Parámetros ===
'excelPath = "C:\ReporteRegulatorioRpa\Temp\Refacturacion_regular_v2.xlsm"
hojaNombre = "TC"                   ' Nombre de la hoja


' === Inicializa Excel ===
'Genera un objeto de tipo Excel Application
Set objExcel = CreateObject("Excel.Application")

'Parámetro para indicar si se quiere visible la aplicación de Excel
objExcel.Application.Visible = False
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = False
'Parámetro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False

Set objWorkbook = objExcel.Workbooks.Open(excelPath, 0)
Set objSheet = objWorkbook.Sheets(hojaNombre)

valorBusqueda = "Información diaria"                 ' Valor a buscar

' Ultima fila con datos
ultimaFila = objSheet.Cells(objSheet.Rows.Count, 1).End(-4162).Row
' Obtener valor de la columna A de la última fila
valorUltimaFila = objSheet.Cells(ultimaFila, 1).Value


' === Guardar y cerrar ===
objWorkbook.Save
objWorkbook.Close False
objExcel.Quit

'Devuelve el error en caso de
If Err.Number <> 0 Then
    ' Guardar y cerrar el libro de refacturación
    objWorkbook.Save
    objWorkbook.Close
    ' Cerrar la aplicación de Excel
    objExcel.Quit
    Msg = "Error was generated by " & Err.Source & ". " & Err.Description
    WScript.StdOut.WriteLine Msg
Else
    WScript.StdOut.WriteLine valorUltimaFila & ":Script executed successfully."
End if
