VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Productos 
   Caption         =   "Form2"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5775
   LinkTopic       =   "Form2"
   ScaleHeight     =   8250
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtdescripcion 
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
      Left            =   2760
      TabIndex        =   17
      Top             =   3960
      Width           =   2535
   End
   Begin VB.TextBox txtproducto 
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
      TabIndex        =   15
      Top             =   1080
      Width           =   2775
   End
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
      Left            =   2160
      TabIndex        =   13
      Top             =   6960
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo dc1 
      Height          =   480
      Left            =   1440
      TabIndex        =   11
      Top             =   5400
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   847
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtprecio 
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1800
      TabIndex        =   9
      Top             =   4650
      Width           =   3495
   End
   Begin VB.TextBox txttamaño 
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
      Left            =   2040
      TabIndex        =   8
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox txtstock 
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1680
      TabIndex        =   7
      Top             =   3240
      Width           =   3615
   End
   Begin VB.TextBox txtcolor 
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1680
      TabIndex        =   6
      Top             =   2520
      Width           =   3615
   End
   Begin VB.CommandButton guardar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Guardar Producto"
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
      Left            =   1920
      MaskColor       =   &H00FFC0FF&
      TabIndex        =   5
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   360
      TabIndex        =   16
      Top             =   4080
      Width           =   2370
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   360
      TabIndex        =   14
      Top             =   1200
      Width           =   2070
   End
   Begin VB.Label lbltipo 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   8040
      TabIndex        =   12
      Top             =   600
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   360
      TabIndex        =   10
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tamaño:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   1605
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   1275
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   360
      TabIndex        =   2
      Top             =   4800
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   360
      TabIndex        =   1
      Top             =   3360
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Nuevo producto"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5100
   End
End
Attribute VB_Name = "Productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub dc1_Change()
    With RsTipoProducto
        .Requery
        .Find "Tipo='" & Trim(dc1.Text) & "'"
        If .BOF Or .EOF Then Exit Sub
        lbltipo.Caption = !Id_Tipoproducto
    End With
End Sub


Private Sub Form_Load()
    Productos.Picture = LoadPicture(App.Path & "\IMG\Fondo2.jpg")
    Producto
    TipoProducto
    Set dc1.RowSource = RsTipoProducto
    dc1.BoundColumn = "Tipo"
    dc1.ListField = "Tipo"
End Sub
Private Sub guardar_Click()
    If txtproducto.Text = "" Then MsgBox "Ingrese el NOMBRE del producto", vbInformation, "Aviso": txtproducto.SetFocus: Exit Sub
    If Txtcolor.Text = "" Then MsgBox "Ingrese el COLOR del producto", vbInformation, "Aviso": Txtcolor.SetFocus: Exit Sub
    If Val(Txtstock.Text) = 0 Then MsgBox "Ingrese el STOCK del producto", vbInformation, "Aviso": Txtstock.SetFocus: Exit Sub
    If Txttamaño.Text = "" Then MsgBox "Ingrese el TAMAÑO del producto", vbInformation, "Aviso": Txttamaño.SetFocus: Exit Sub
    If txtdescripcion.Text = "" Then MsgBox "Ingrese la DESCRIPCIÓN del producto", vbInformation, "Aviso": Txttamaño.SetFocus: Exit Sub
    If Val(txtprecio.Text) = 0 Then MsgBox "Ingrese el PRECIO del producto", vbInformation, "Aviso": txtprecio.SetFocus: Exit Sub
    If dc1.Text = "" Then MsgBox "Seleccione el TIPO del producto", vbInformation, "Aviso": dc1.SetFocus: Exit Sub
    With RsProductos
        .Requery
        .AddNew
        !Producto = txtproducto.Text
        !Color = Txtcolor.Text
        !Stock = Txtstock.Text
        !Precio = txtprecio.Text
        !Tamaño = Txttamaño.Text
        !Tipo = dc1.Text
        !Descripcion = txtdescripcion.Text
        !Id_Tipoproducto = lbltipo.Caption
        .UpdateBatch
End With
MsgBox "El producto ha sido registrado correctamente", vbInformation, "Aviso"
limpiar
End Sub

Sub limpiar()
    txtproducto.Text = ""
    Txtcolor.Text = ""
    Txtstock.Text = ""
    txtprecio.Text = ""
    Txttamaño.Text = ""
    txtdescripcion.Text = ""
    dc1.Text = ""
End Sub

