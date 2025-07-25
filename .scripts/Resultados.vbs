On Error Resume Next

Set objArgs = WScript.Arguments

rutaArchivoTxt = objArgs(0)
rutaExcelDestino = objArgs(1)

' Ruta del archivo txt
'rutaArchivoTxt = "C:\RPA_Process\Timbrado\XML\resultados_2025716.txt"
' Ruta del archivo Excel EXISTENTE
'rutaExcelDestino = "C:\RPA_Process\Timbrado\Refacturacion_regular_v2.xlsm"

'Genera un objeto de tipo Excel Application
Set objExcel = CreateObject("Excel.Application")

'Parámetro para indicar si se quiere visible la aplicación de Excel
objExcel.Application.Visible = False
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = False
'Parámetro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False

' Abrir libro existente
Set objWorkbook = objExcel.Workbooks.Open(rutaExcelDestino, 0)
Set objHoja = objWorkbook.Worksheets("Resultado")

' Limpiar contenido anterior en la hoja
objHoja.Cells.ClearContents

' Abrir archivo de texto
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objArchivo = objFSO.OpenTextFile(rutaArchivoTxt, 1, False, 0) ' 0 = ANSII

fila = 1

Do Until objArchivo.AtEndOfStream
    line = objArchivo.ReadLine
    campos = Split(line, vbTab)

    Dim col
    For col = 0 To UBound(campos)
        objHoja.Cells(fila, col + 1).Value = campos(col)
    Next
    fila = fila + 1
Loop

' Borrar fila 1
objHoja.Rows(1).Delete


objArchivo.Close

' Ajustar ancho de columnas automáticamente
objHoja.Columns("A:G").AutoFit

' Guardar y cerrar
objWorkbook.Save
objWorkbook.Close False
objExcel.Quit

If Err.Number <> 0 Then
    Msg = "Error was generated by " & Err.Source & ". " & Err.Description
    WScript.StdOut.WriteLine Msg
Else
    WScript.StdOut.WriteLine "Script executed successfully."
End If