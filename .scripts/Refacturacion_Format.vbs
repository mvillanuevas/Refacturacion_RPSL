On Error Resume Next

Set objArgs = WScript.Arguments

WorkbookPathRexmex = objArgs(0)
WorkbookPathRef = objArgs(1)
ActualMonth = objArgs(2)
saveLastRow = objArgs(3)
WorkbookPathLayout = objArgs(4)

'WorkbookPathRexmex = "C:\Users\se109874\OneDrive - Repsol\Documentos\Refacturacion\REXMEX - Cuenta Operativa 2025_120525.xlsx"
'WorkbookPathRef = "C:\Users\se109874\OneDrive - Repsol\Documentos\Refacturacion\Layout refacturación may-25.xlsx"
'ActualMonth = "3"
'saveLastRow = "843"

WorkbookSheetRexmex = "Cuenta Operativa"
WorkbookSheetLayout = "Layout"

'Genera un objeto de tipo Excel Application
Set objExcel = CreateObject("Excel.Application")

'Parámetro para indicar si se quiere visible la aplicación de Excel
objExcel.Application.Visible = False
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = False
'Parámetro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False

'Abre libro Excel
Set objWorkbookPathRef = objExcel.Workbooks.Open(WorkbookPathRef, 0)
Set objWorkbookSheetRefL = objWorkbookPathRef.Worksheets(WorkbookSheetLayout)

saveLastRow = CInt(saveLastRow)
' Encontrar la última fila con datos en la hoja de Layout refacturación
LastRow = objWorkbookSheetRefL.Cells(objWorkbookSheetRefL.Rows.Count, 4).End(-4162).Row

' Aplicar negritas a un rango específico
With objWorkbookSheetRefL.Range("A" & saveLastRow & ":B" & LastRow)
    .Font.Bold = True
End With

' Aplicar All borders a un rago específico
With objWorkbookSheetRefL.Range("C" & saveLastRow & ":AF" & LastRow)
    .Borders.LineStyle = 1 ' xlContinuous
    .Borders.Weight = 2 ' xlMedium
End With

' Crear un arreglo de columnas para aplicar el formato Right border
Dim rightBorderCols
rightBorderCols = Array("C", "D", "E", "K", "M", "N", "V", "W", "Z", "AC", "AF")
' Aplicar Right border a las columnas especificadas
Dim col
For Each col In rightBorderCols
    With objWorkbookSheetRefL.Range(col & saveLastRow & ":" & col & LastRow)
        .Borders(10).LineStyle = 1 ' xlContinuous
        .Borders(10).Weight = -4138 ' xlMedium
    End With
Next
' Aplicar color de fondo a un rango específico
With objWorkbookSheetRefL.Range("C" & saveLastRow & ":AF" & LastRow)
    .Interior.Color = RGB(217, 225, 242) ' Color azul claro
End With

' Crear un nombre de hoja basado en la fecha y hora actual
sheetName = "Layout " & Day(Now) & Month(Now) & Year(Now) & "_" & Second(Now)

' Crear una copia de la hoja actual sobre el mismo libro
If Not SheetExists(objWorkbookPathRef, sheetName) Then
    objWorkbookSheetRefL.Copy objWorkbookPathRef.Sheets(objWorkbookPathRef.Sheets.Count)
    Set objSheetCopy = objWorkbookPathRef.Sheets(objWorkbookPathRef.Sheets.Count - 1)
    objSheetCopy.Name = sheetName
Else
    objWorkbookPathRef.Sheets(sheetName).Delete
    objWorkbookSheetRefL.Copy objWorkbookPathRef.Sheets(objWorkbookPathRef.Sheets.Count)
    Set objSheetCopy = objWorkbookPathRef.Sheets(objWorkbookPathRef.Sheets.Count - 1)
    objSheetCopy.Name = sheetName
End If

Set objWorkbookSheetRefLN = objWorkbookPathRef.Worksheets(sheetName)

' Copiar y pegar como valores todas las celdas de una hoja
With objWorkbookSheetRefLN.UsedRange
    .Copy
    .PasteSpecial -4163 ' xlPasteValues
End With
objExcel.CutCopyMode = False

' Por cada celda en la columna E, buscar si contiene "pep" y eliminar el texto a la derecha a partir de "pep"
Dim cell
For Each cell In objWorkbookSheetRefLN.Range("E" & saveLastRow & ":E" & LastRow)
    If Not IsEmpty(cell.Value) And Not IsNull(cell.Value) Then
        If InStr(1, cell.Value, "pep", vbTextCompare) > 0 Then
            posicion = InStr(1, cell.Value, "pep", vbTextCompare) - 1
            cell.Value = Left(cell.Value, posicion)
        End If
    End If
Next

refacturacionSheets = Array("BL29", "BL10", "BL11", "BL14")
' Iteraer sobre las hojas de refacturación y eliminar el contenido desde la fila 2 hasta la última fila
Dim i

For i = LBound(refacturacionSheets) To UBound(refacturacionSheets)
    Dim sheetName
    sheetName = refacturacionSheets(i)
    
    ' Verificar si la hoja existe en el libro de refacturación
    If SheetExists(objWorkbookPathRef, sheetName) Then
        Set objWorkbookSheetRef = objWorkbookPathRef.Worksheets(sheetName)
        
        ' Encontrar la última fila con datos en la hoja de Layout refacturación
        lastRow = objWorkbookSheetRef.Cells(objWorkbookSheetRef.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp
        
        ' Limpiar el contenido desde la fila 2 hasta la última fila
        If lastRow > 1 Then
            objWorkbookSheetRef.Range("A2:AV" & lastRow).ClearContents
        End If
        
        ' Quitar el modo de corte/copia
        objExcel.CutCopyMode = False
    End If
Next
' Guardar con otro nombre el libro de refacturación
objWorkbookPathRef.SaveAs WorkbookPathLayout, 51 ' 51 = xlOpenXMLWorkbook (xlsx)
objWorkbookPathRef.Close
' Cerrar la aplicación de Excel
objExcel.Quit

'Devuelve el error en caso de
If Err.Number <> 0 Then
    ' Cerrar la aplicación de Excel
    objExcel.Quit
    Msg = "Error was generated by " & Err.Source & ". " & Err.Description
    WScript.StdOut.WriteLine Msg
Else
    WScript.StdOut.WriteLine "Script executed successfully."
End if

'____________________________________________________________________________________________________________________________________________
' Función para validar si una hoja existe en un libro de Excel
Function SheetExists(wb, sheetName)
    Dim s
    SheetExists = False
    For Each s In wb.Sheets
        If StrComp(s.Name, sheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next
End Function