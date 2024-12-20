VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form3 
   BackColor       =   &H80000016&
   Caption         =   "Form3"
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
   LinkTopic       =   "Form3"
   ScaleHeight     =   6675
   ScaleWidth      =   14475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   8040
      TabIndex        =   13
      Top             =   3240
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      _Version        =   393216
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
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "Inserta Imagen"
      Height          =   405
      Left            =   12240
      TabIndex        =   12
      Top             =   1080
      Width           =   1830
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11520
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      TabIndex        =   11
      Top             =   5880
      Width           =   6615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FF80&
      Caption         =   "Añadir Elemento"
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
      TabIndex        =   10
      Top             =   5880
      Width           =   6615
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
      TabIndex        =   9
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox Text1 
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
      TabIndex        =   3
      Top             =   1800
      Width           =   6015
   End
   Begin VB.TextBox Text2 
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
   Begin VB.TextBox Text4 
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
      TabIndex        =   1
      Top             =   3960
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   4335
      Left            =   360
      Top             =   1080
      Width           =   4935
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
      TabIndex        =   8
      Top             =   1080
      Width           =   2175
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
      TabIndex        =   7
      Top             =   1800
      Width           =   2145
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
      TabIndex        =   6
      Top             =   3960
      Width           =   2100
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
      TabIndex        =   5
      Top             =   3240
      Width           =   2115
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
      TabIndex        =   4
      Top             =   2520
      Width           =   2145
   End
   Begin VB.Label lblMiVision 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Añadir Elemento"
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
      Width           =   14175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variables Globales
Dim imageName As String

' Funcion para subir la imagen al sistema
Private Sub cmdCommand1_Click()
    
    ' Declarar la variable para la ruta de la imagen
    Dim rutaimagen As String
    
    ' Se configura el espacio para mostrar el directorio de archivo para seleccionar la imagen deseada
    ' Solo se aceptan archvio en formato JPG
    With CommonDialog1
        .DialogTitle = "Selecciona las imágenes"
        .Filter = "Archivos de imagen | *.jpg"
        .ShowOpen
       rutaimagen = .FileName
    End With
    
    ' Verificar si se ha seleccionado un archivo (no vacío)
    If rutaimagen <> "" Then
    
        ' Cargar la imagen seleccionada en el control Image1
        Set Image1.Picture = LoadPicture(rutaimagen)
        
        ' Ajustar el tamaño de la imagen
        Image1.Stretch = True
        Image1.Width = 4935
        Image1.Height = 4335
                   
        ' Se obtiene el nombre de la imagen y se muestra en su sección correspondiente
        imageName = Mid(rutaimagen, InStrRev(rutaimagen, "\") + 1)
        Text5.Text = imageName
        
        ' Se establece la ruta para guardar una copia de la imagen
        savePath = App.Path & "\imagenes\" & imageName

        ' Se guarda la imagen en la nueva ruta
        SavePicture Image1.Picture, savePath
    
    Else
        
        ' Si no se seleccionó la imagen, se muestra una imagen
        MsgBox "No se seleccionó ninguna imagen.", vbExclamation
        
    End If
    
End Sub

' Función para regresar a la pantalla principal
Private Sub Command3_Click()
    Form1.Show
    Unload Me
End Sub

' Funcion para guardar el nuevo proposito en la base de datos
Private Sub Command4_Click()

    ' Manejo de errores
    On Error GoTo ErrorHandler
    
    ' Conexión con la base de datos
    Dim conn As Object
    Set conn = ObtenerConexion()
    
    ' Validador para evitar datos vacio. Termina la función si no hay titulo y/o imagen
    If Text1.Text = "" Or imageName = "" Then
        
        MsgBox "Ingrese un elemento valido"
        Exit Sub
        
    End If
    
    ' Recordset para ejecutar una consulta validadora en SQL
    Dim rs_validador As Object
    Set rs_validador = CreateObject("ADODB.Recordset")
        
    ' Se diseña y ejecuta la consulta SQL. Además, el resultado se asigna a una variable para su validación
    Dim validador As String
    Dim repeticiones As Integer
    validador = "SELECT COUNT(*) as valor FROM propositos WHERE titulo = '" & Replace(Text1.Text, "'", "''") & "';"
    rs_validador.Open validador, conn
    repeticiones = rs_validador.Fields("valor").Value

    ' Se valida que no existan titulos repetidos
    If repeticiones = 0 Then
    
       ' Se adapta la fecha a formato SQL Server
       Dim fecha As String
       fecha = Format(DTPicker1.Value, "yyyy-mm-dd")
       
       ' Se crea el recordset para ejecutar la consulta principal en SQL
       Dim rs As Object
       Set rs = CreateObject("ADODB.Recordset")
       
       ' Se diseña y ejecuta la consulta SQL
       Dim sql As String
       sql = "INSERT INTO propositos (imagen, titulo, categoria, fecha_terminacion, descripcion) " & _
             "VALUES ('" & imageName & "', '" & Text1.Text & "', '" & Text2.Text & "', '" & fecha & "', '" & Text4.Text & "');"
       conn.Execute sql
       
       ' Se notifica al usuario que la consulta se ejecuto y se redirije a la pantalla principal
       MsgBox "Propósito agregado exitosamente"
       Form1.Show
       Unload Me
       
    Else
       
       ' Si ya existe un titulo similar, se muestra este mensaje
       MsgBox "Ya existe un proposito con ese titulo. Ingrese un titulo nuevo"
       Exit Sub
       
    End If
    
    ' Se cierra recordset y conexión
    rs.Close
    Set rs = Nothing
    rs_validador.Close
    Set rs_validador = Nothing
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

