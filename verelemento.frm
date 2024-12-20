VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   BackColor       =   &H80000016&
   Caption         =   "Form2"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   6675
   ScaleMode       =   0  'User
   ScaleWidth      =   6163.025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   8040
      TabIndex        =   12
      Top             =   3240
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   179961857
      CurrentDate     =   45645
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8040
      TabIndex        =   11
      Top             =   1080
      Width           =   6015
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Regresar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5880
      Width           =   6660
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Borrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   6615
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   8040
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3960
      Width           =   6015
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8040
      TabIndex        =   2
      Top             =   2520
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8040
      TabIndex        =   1
      Top             =   1800
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   4335
      Left            =   360
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Label lblMiVision 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ver Elemento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   14175
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Categoría"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5640
      TabIndex        =   9
      Top             =   2520
      Width           =   2145
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Fecha de terminación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5640
      TabIndex        =   8
      Top             =   3240
      Width           =   2115
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5640
      TabIndex        =   7
      Top             =   3960
      Width           =   2100
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Título"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5640
      TabIndex        =   6
      Top             =   1800
      Width           =   2145
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Imagen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5640
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variables Globales
Public proposito As String

' Función para eliminar un proposito de la lista
Private Sub Command1_Click()
    
    ' Manejo de errores
    On Error GoTo ErrorHandler
    
    ' Conexión con la base de datos
    Dim conn As Object
    Set conn = ObtenerConexion()
    
    ' Recordset para ejecutar la consulta SQL
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Se crea una variable para guardar la respuesta del usuario sobre eliminar el proposito seleccionado
    Dim respuesta As Integer
    respuesta = MsgBox("¿Estás seguro de que quieres eliminar este propósito?", vbYesNo + vbQuestion, "Confirmación")

    ' Condicion para determinar la accion a realizar de acuerdo al valor de respuesta
    If respuesta = vbYes Then
        
        ' Al aceptar el borrado, se diseña y ejecuta la consulta SQL
        Dim sql As String
        sql = "DELETE FROM propositos WHERE titulo = '" & proposito & "';"
        rs.Open sql, conn
        
        ' Se muestra notificación al usuario y se redirije a la pantalla principal
        MsgBox "Proposito eliminado exitosamente"
        Form1.Show
        Unload Me
        
    Else
        
        ' Si la respuesta es NO, solo se sale de la función
        Exit Sub
        
    End If
    
    ' Se cierra recordset y conexión
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Exit Sub

ErrorHandler:

    ' Se maneja el error
    MsgBox "Ha ocurrido un error: " & Err.Description, vbCritical, "Error"
    
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    
End Sub

' Función para regresar a la pantalla principal
Private Sub Command3_Click()
    Form1.Show
    Unload Me
End Sub

' Funcion principal del formulario
Private Sub Form_Load()
    
    ' Manejo de errores
    On Error GoTo ErrorHandler
    
    ' Se recibe el datos del titulo que se quiere visualizar
    Text1.Text = proposito
    
    ' Conexión con la base de datos
    Dim conn As Object
    Set conn = ObtenerConexion()
    
    ' Recordset para ejecutar la consulta SQL
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Se diseña y ejecuta la consulta SQL
    Dim sql As String
    sql = "SELECT TOP 1 * FROM propositos WHERE titulo = '" & proposito & "';"
    rs.Open sql, conn
    
    ' Se asignan los resultados obtenidos a sus espacios correspondientes
    Text1.Text = rs.Fields("titulo").Value
    Text2.Text = rs.Fields("categoria").Value
    DTPicker1.Value = rs.Fields("fecha_terminacion").Value
    Text4.Text = rs.Fields("descripcion").Value
    Text5.Text = rs.Fields("imagen").Value
    
    ' Se obtiene la ruta de la imagen y se muestra en su espacio
    Dim imagePath As String
    imagePath = App.Path & "\imagenes\" & rs.Fields("imagen").Value
    Image1.Picture = LoadPicture(imagePath)
    
    ' Se ajusta el tamaño de la imagen
    Image1.Stretch = True
    Image1.Width = 2101.176
    Image1.Height = 4335
    
    ' Se cierra recordset y conexión
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Exit Sub

ErrorHandler:

    ' Se maneja el error
    MsgBox "Ha ocurrido un error: " & Err.Description, vbCritical, "Error"
    
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    
End Sub
