Attribute VB_Name = "Module1"
' M�dulo: ModuloConexion
Public Function ObtenerConexion() As Object
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Cadena de conexi�n
    Dim connectionString As String
    connectionString = "Provider=SQLOLEDB;Data Source=TU_SERVIDOR;Initial Catalog=TU_BASEDATOS;Integrated Security=SSPI;"
    
    ' Establecer la conexi�n
    conn.Open connectionString
    
    ' Devolver el objeto de conexi�n
    Set ObtenerConexion = conn
End Function

