Imports System.Data
Imports System.Reflection

' Cargar el ensamblado
Dim asm = Assembly.LoadFrom("C:\sqlite3\System.Data.SQLite.dll")
Dim connType = asm.GetType("System.Data.SQLite.SQLiteConnection")
Dim cmdType = asm.GetType("System.Data.SQLite.SQLiteCommand")
Dim rdrType = asm.GetType("System.Data.SQLite.SQLiteDataReader")
Dim cmdBehaviorType = GetType(CommandBehavior)

' Crear la conexión
Dim conn = Activator.CreateInstance(connType, "Data Source=C:\sqlite3\ReporteRegulatorio.db;Version=3;")
connType.GetMethod("Open").Invoke(conn, Nothing)

' Crear y ejecutar el comando
Dim cmd = Activator.CreateInstance(cmdType, New Object() {"SELECT MAX(Folio) FROM Folios", conn})
Dim execReaderMethod = cmdType.GetMethod("ExecuteReader", New Type() {cmdBehaviorType})
Dim rdr = execReaderMethod.Invoke(cmd, New Object() {CommandBehavior.Default})

' Cargar los resultados en un DataTable
Dim dt As New DataTable()
dt.Load(CType(rdr, IDataReader))

' Cerrar el lector
rdrType.GetMethod("Close").Invoke(rdr, Nothing)

' Liberar el comando (opcional)
' CType(cmd, IDisposable).Dispose() ' <- más correcto que buscar Close por Reflection

' Obtener el resultado
Dim out_Folio As Integer = -1
If dt.Rows.Count > 0 AndAlso Not IsDBNull(dt.Rows(0)(0)) Then
    out_Folio = Convert.ToInt32(dt.Rows(0)(0))
End If