VERSION 5.00
Begin VB.Form Productos 
   Caption         =   "Form2"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4560
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txttipo 
      Height          =   285
      Left            =   6240
      TabIndex        =   11
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Txtprecio 
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Txttamaño 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtstock 
      Height          =   285
      Left            =   6240
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Txtcolor 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton guardar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "guardar producto"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      MaskColor       =   &H00FFC0FF&
      TabIndex        =   5
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "tamaño"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "color:"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "precio"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "stock"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Nuevo producto"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Producto
With RsProductos
If .EOF Or .BOF Then Exit Sub
End With
End Sub
Private Sub guardar_Click()
With RsProductos
    .Requery
    .AddNew
    !Color = Txtcolor.Text
    !stock = txtstock.Text
    !Precio = Txtprecio.Text
    !Tipo = Txttipo.Text
    !Tamaño = Txttamaño.Text
End With
End Sub

