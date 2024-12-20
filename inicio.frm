VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18585
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   18585
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   5310
      Left            =   1440
      TabIndex        =   3
      Top             =   1440
      Width           =   4815
   End
   Begin VB.CommandButton cmdX 
      BackColor       =   &H008080FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdCommand1 
      BackColor       =   &H0080FF80&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   17280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   6600
      Picture         =   "inicio.frx":0000
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   5280
   End
   Begin VB.Image ImagenPrincipal 
      Height          =   5295
      Left            =   12240
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Line Line17 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   12000
      X2              =   12000
      Y1              =   1200
      Y2              =   6960
   End
   Begin VB.Line Line16 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   6480
      X2              =   6480
      Y1              =   1200
      Y2              =   6960
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   12000
      X2              =   12000
      Y1              =   6960
      Y2              =   7320
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   5775
      Index           =   0
      Left            =   960
      Top             =   1200
      Width           =   5535
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   6480
      X2              =   6480
      Y1              =   6960
      Y2              =   7320
   End
   Begin VB.Line Line13 
      X1              =   17760
      X2              =   720
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line12 
      X1              =   17640
      X2              =   840
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line11 
      X1              =   17760
      X2              =   17760
      Y1              =   1440
      Y2              =   7200
   End
   Begin VB.Line Line10 
      X1              =   17640
      X2              =   17640
      Y1              =   1320
      Y2              =   7080
   End
   Begin VB.Line Line9 
      X1              =   840
      X2              =   840
      Y1              =   1320
      Y2              =   7080
   End
   Begin VB.Line Line8 
      X1              =   720
      X2              =   720
      Y1              =   1560
      Y2              =   7200
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   17880
      X2              =   17880
      Y1              =   1560
      Y2              =   7320
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   600
      X2              =   600
      Y1              =   1560
      Y2              =   7320
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   600
      X2              =   17880
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   960
      X2              =   600
      Y1              =   6960
      Y2              =   7320
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   17520
      X2              =   17880
      Y1              =   6960
      Y2              =   7320
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   17520
      X2              =   17880
      Y1              =   1200
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   960
      X2              =   600
      Y1              =   1200
      Y2              =   1560
   End
   Begin VB.Label lblLabel1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mi Vision Board"
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
      TabIndex        =   0
      Top             =   240
      Width           =   18255
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000018&
      FillColor       =   &H000000FF&
      Height          =   5415
      Left            =   17520
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000018&
      FillColor       =   &H000000FF&
      Height          =   375
      Left            =   17520
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000018&
      FillColor       =   &H000000FF&
      Height          =   5415
      Left            =   600
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000018&
      FillColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   375
      Left            =   600
      Top             =   6960
      Width           =   17295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   5775
      Index           =   1
      Left            =   6480
      Top             =   1200
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   5775
      Index           =   2
      Left            =   12000
      Top             =   1200
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Se dirige a la sección para crear un nuevo proposito
Private Sub cmdCommand1_Click()
    Form3.Show
    Unload Me
End Sub

' Termina el programa
Private Sub cmdX_Click()
    Unload Me
End Sub

' Funcion principal
Private Sub Form_Load()
    
    ' Manejo de errores
    On Error GoTo ErrorHandler
    
    ' Conexión con la base de datos
    Dim conn As Object
    Set conn = ObtenerConexion()
    
    ' Recordset para ejecutar la consulta SQL
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Se diseña y ejecuta la consulta SQL
    Dim sql As String
    sql = "SELECT titulo FROM propositos;"
    rs.Open sql, conn
    
    ' Verificar si el Recordset contiene datos
    If rs.EOF Then
        MsgBox "No se encontraron registros.", vbExclamation
        Exit Sub
    End If
    
    ' Añade los titulos de los propositos a la lista
    Do Until rs.EOF
        List1.AddItem rs.Fields("titulo").Value
        rs.MoveNext
    Loop
    
    ' Ajusta el tamaño de la imagen
    ImagenPrincipal.Stretch = True
    ImagenPrincipal.Width = 4815
    ImagenPrincipal.Height = 5295

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

' Sección de imagenes
Private Sub ImagenPrincipal_Click()

    ' Condicion que valida que exista propositos para mostrar
    If List1.Text = "" Then
        MsgBox "Seleccione un proposito para visualizar"
    Else
        Form4.proposito = List1.Text
        Form4.Show
        Unload Me
    End If
    
End Sub

'Lista de propositos
Private Sub List1_Click()
    
    ' Manejo de errores
    On Error GoTo ErrorHandler
     
    ' Conexión con la base de datos
    Dim conn As Object
    Set conn = ObtenerConexion()
    
    ' Recordset para ejecutar la consulta SQL
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Se diseña y ejecuta la consulta SQL
    Dim sql As String
    sql = "SELECT TOP 1 imagen FROM propositos WHERE titulo = '" & Replace(List1.Text, "'", "''") & "';"
    rs.Open sql, conn
    
    ' Se busca la imagen y se muestra en el espacio correspondiente
    Dim imagePath As String
    imagePath = App.Path & "\imagenes\" & rs.Fields("imagen").Value
    ImagenPrincipal.Picture = LoadPicture(imagePath)
    
    ' Se cierra recordset y conexión
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Exit Sub ' Asegurarse de que el flujo normal salga antes de ejecutar el manejador de errores

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

