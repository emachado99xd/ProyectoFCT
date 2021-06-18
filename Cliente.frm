VERSION 5.00
Begin VB.Form Cliente 
   Caption         =   "Form2"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6180
   LinkTopic       =   "Form2"
   ScaleHeight     =   6435
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   13
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txttelefono 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox txtdireccion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   10
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox txtapellido 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox txtcedula 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox txtnombre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton cmdguardar 
      Caption         =   "Guardar Cliente"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   5
      Top             =   4560
      Width           =   1695
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
      Left            =   10800
      TabIndex        =   12
      Top             =   2280
      Width           =   120
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nuevo cliente"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   1080
      TabIndex        =   6
      Top             =   120
      Width           =   4260
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   720
      TabIndex        =   4
      Top             =   3840
      Width           =   2100
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "APELLIDO:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   720
      TabIndex        =   2
      Top             =   2400
      Width           =   2010
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cédula:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   720
      TabIndex        =   1
      Top             =   3120
      Width           =   1560
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   1695
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
    Cliente.Picture = LoadPicture(App.Path & "\IMG\Fondo2.jpg")
End Sub

Sub limpiar()
    txtnombre.Text = ""
    txtapellido.Text = ""
    txtcedula.Text = ""
    txtdireccion.Text = ""
    txttelefono.Text = ""
    
End Sub
