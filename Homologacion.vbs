On Error Resume Next

Set objArgs = WScript.Arguments

excelFilePath = objArgs(0)
wsNames = objArgs(1)
colNames = objArgs(2)
row = objArgs(3)

' Cambia la ruta del archivo Excel según corresponda
'excelFilePath = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\03 IMSS.xlsx"
'wsNames = "Sheet1"
'colNames = "Cuenta|Nombre del usuario|Sociedad|Documento compras|Clase de cuenta|Indicador Debe/Haber|Clave contabiliz.|Cuenta de mayor|Sociedad GL asociada|Acreedor|Nombre 1|Ejercicio / mes|Ind.cesión créditos|Indicador impuestos|Nº documento|Clase de documento|Fe.contabilización|Fecha de documento|Texto|Asignación|Referencia|Joint Venture|Concepto|Centro de coste|Stat.part.abiertas/compens.|Tipo de coste|Importe en moneda doc.|Moneda del documento|Importe en moneda local|Moneda local|Importe en ML2|Mon.local 2|Importe en ML3|Moneda local 3|Cta.contrapartida"
'row = "1"

'Genera un objeto de tipo Excel Application
Set objExcel = CreateObject("Excel.Application")

'Parámetro para indicar si se quiere visible la aplicación de Excel
objExcel.Application.Visible = False
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = False
'Parámetro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False


Set objWorkbook = objExcel.Workbooks.Open(excelFilePath, 0)

wsNames = Split(wsNames, "|")
colNames = Split(colNames, "|")
row = CInt(row)

result = ValidarHojasYColumnas(objWorkbook, wsNames, colNames, row)

objWorkbook.Close False
objExcel.Quit

If result = "" Then
    WScript.StdOut.WriteLine "Script executed successfully."
Else
    WScript.StdOut.WriteLine "Error, no existe: " & result
End If

'________________________________________________________________________________________
' Function para validar la existencia de hojas y columnas en un libro de Excel
' ' Parámetros:
'   objWorkbook: Objeto Workbook de Excel
'   wsNames: Array de nombres de hojas a validar
'   colNames: Array de nombres de columnas a validar
'   row: Fila en la que se deben buscar las columnas
' ' Retorna: Un string con los nombres de hojas y columnas no encontradas, separados
'   por comas. Si todo está correcto, retorna un string vacío.
'________________________________________________________________________________________
Function ValidarHojasYColumnas(objWorkbook, wsNames, colNames, row)
    ' Asignar variables
    Dim i, j, k, found, wsNotFound, colFound, colNotFound, objWorksheet
    Dim lastCol, col, msg

    wsNotFound = ""
    msg = ""

    ' Validar si el libro tiene hojas
    For i = 0 To UBound(wsNames)
        found = False
        ' Buscar si la hoja existe en el libro
        For Each objWorksheet In objWorkbook.Worksheets
            If objWorksheet.Name = wsNames(i) Then
                found = True
                Exit For
            End If
        Next
        ' Si la hoja no se encontró, agregar al mensaje
        If Not found Then
            wsNotFound = wsNotFound & "'" & wsNames(i) & "',"
        Else
            ' Validar columnas en la fila indicada
            lastCol = objWorksheet.Cells(row, objWorksheet.Columns.Count).End(-4159).Column  ' -4159 = xlToLeft
            colNotFound = ""
            ' Buscar cada columna en la fila especificada
            For j = 0 To UBound(colNames)
                colFound = False
                ' Verificar si la columna existe en la fila indicada
                For col = 1 To lastCol
                    If Trim(LCase(objWorksheet.Cells(row, col).Value)) = LCase(colNames(j)) Then
                        colFound = True
                        Exit For
                    End If
                Next
                ' Si la columna no se encontró, agregar al mensaje
                If Not colFound Then
                    colNotFound = colNotFound & "'" & colNames(j) & "', "
                End If
            Next
            ' Si hay columnas no encontradas, agregar al mensaje
            If colNotFound <> "" Then
                ' Quitar última coma y espacio
                If Right(colNotFound,2) = ", " Then colNotFound = Left(colNotFound, Len(colNotFound)-2)
                    msg = msg & "'" & colNotFound & "',"
                End If
        End If
    Next
    ' Retornar mensaje de error si hay hojas o columnas no encontradas
    ValidarHojasYColumnas = "Hoja: " & wsNotFound & " Columna: " & msg
End Function