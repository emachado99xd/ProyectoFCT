VERSION 5.00
Begin VB.Form Inicio 
   Caption         =   "Form2"
   ClientHeight    =   4125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6795
   LinkTopic       =   "Form2"
   ScaleHeight     =   4125
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Facturas"
      BeginProperty Font 
         Name            =   "Unispace"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Control de productos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3720
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Control de clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Facturación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   1
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "¿Qué acción desea realizar?"
      BeginProperty Font 
         Name            =   "Unispace"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.Show
End Sub





Private Sub Command2_Click()
    Anular.Show
End Sub

Private Sub Command4_Click()
    ModificarCliente.Show
End Sub

Private Sub Command5_Click()
    ModificarProducto.Show
End Sub
