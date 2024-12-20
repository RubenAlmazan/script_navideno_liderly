Attribute VB_Name = "Module1"
' Módulo: ModuloConexion
Public Function ObtenerConexion() As Object
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Cadena de conexión
    Dim connectionString As String
    connectionString = "Provider=SQLOLEDB;Data Source=TU_SERVIDOR;Initial Catalog=TU_BASEDATOS;Integrated Security=SSPI;"
    
    ' Establecer la conexión
    conn.Open connectionString
    
    ' Devolver el objeto de conexión
    Set ObtenerConexion = conn
End Function

