VERSION 5.00
Begin VB.Form Cliente 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form2"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8250
   LinkTopic       =   "Form2"
   ScaleHeight     =   4980
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   13
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txttelefono 
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox txtdireccion 
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox txtapellido 
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtcedula 
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtnombre 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton cmdguardar 
      Caption         =   "Guardar Cliente"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      TabIndex        =   5
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label lblidcliente 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   8280
      TabIndex        =   12
      Top             =   1800
      Width           =   120
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "nuevo cliente"
      BeginProperty Font 
         Name            =   "Moonbright Demo"
         Size            =   54.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   1440
      TabIndex        =   6
      Top             =   120
      Width           =   5160
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3960
      TabIndex        =   4
      Top             =   2640
      Width           =   1590
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4200
      TabIndex        =   3
      Top             =   1560
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   2
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cédula:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   1365
   End
End
Attribute VB_Name = "Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdguardar_Click()
    If txtnombre.Text = "" Then MsgBox "Ingrese el NOMBRE del cliente", vbInformation, "Aviso": txtnombre.SetFocus: Exit Sub
    If txtapellido.Text = "" Then MsgBox "Ingrese el APELLIDO del cliente", vbInformation, "Aviso": txtapellido.SetFocus: Exit Sub
    If txtcedula.Text = "" Then MsgBox "Ingrese la CEDULA del cliente", vbInformation, "Aviso": txtcedula.SetFocus: Exit Sub
    If txtdireccion.Text = "" Then MsgBox "Ingrese la DIRECCIÓN del cliente", vbInformation, "Aviso": txtdireccion.SetFocus: Exit Sub
    If txttelefono.Text = "" Then MsgBox "Ingrese el TELÉFONO del cliente", vbInformation, "Aviso": txttelefono.SetFocus: Exit Sub
    With RsCliente
        .Requery
        .AddNew
        !Nombre = txtnombre.Text
        !Apellido = txtapellido.Text
        !Cedula = txtcedula.Text
        !Direccion = txtdireccion.Text
        !Telefono = txttelefono.Text
        .UpdateBatch
    End With
    limpiar
    MsgBox "El Cliente ha sido registrado correctamente", vbInformation, "Aviso"
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Clientes
End Sub

Sub limpiar()
    txtnombre.Text = ""
    txtapellido.Text = ""
    txtcedula.Text = ""
    txtdireccion.Text = ""
    txttelefono.Text = ""
    
End Sub
