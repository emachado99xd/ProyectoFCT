VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Productos 
   Caption         =   "Form2"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11970
   LinkTopic       =   "Form2"
   ScaleHeight     =   4770
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtdescripcion 
      Height          =   495
      Left            =   9600
      TabIndex        =   17
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtproducto 
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   9720
      TabIndex        =   13
      Top             =   720
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo dc1 
      Height          =   315
      Left            =   6000
      TabIndex        =   11
      Top             =   3360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.TextBox txtprecio 
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
   Begin VB.TextBox txttamaño 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtstock 
      Height          =   285
      Left            =   6480
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtcolor 
      Height          =   285
      Left            =   6360
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
      Left            =   4920
      MaskColor       =   &H00FFC0FF&
      TabIndex        =   5
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Descripción:"
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
      Left            =   8280
      TabIndex        =   16
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Producto"
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
      TabIndex        =   14
      Top             =   960
      Width           =   1575
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
      Left            =   4320
      TabIndex        =   10
      Top             =   3360
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
      Left            =   5040
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
      Left            =   4920
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
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

