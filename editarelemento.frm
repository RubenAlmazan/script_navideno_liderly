VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000016&
   Caption         =   "Form2"
   ClientHeight    =   6690
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
   ScaleHeight     =   6690
   ScaleMode       =   0  'User
   ScaleWidth      =   6163.025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text9 
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
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1080
      Width           =   6015
   End
   Begin VB.TextBox Text8 
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
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1800
      Width           =   6015
   End
   Begin VB.TextBox Text7 
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
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2520
      Width           =   6015
   End
   Begin VB.TextBox Text6 
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
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3240
      Width           =   6015
   End
   Begin VB.TextBox Text5 
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
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3960
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
      TabIndex        =   2
      Top             =   5880
      Width           =   6615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FF80&
      Caption         =   "Aceptar Cambios"
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
      TabIndex        =   1
      Top             =   5880
      Width           =   6615
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   2520
      Width           =   2145
   End
   Begin VB.Label lblMiVision 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Editar Elemento"
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
      Left            =   135
      TabIndex        =   0
      Top             =   240
      Width           =   14145
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
    Form1.Show
    Unload Me
End Sub
