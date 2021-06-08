VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Productos 
   BackColor       =   &H80000003&
   Caption         =   "Form2"
   ClientHeight    =   5280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9690
   LinkTopic       =   "Form2"
   ScaleHeight     =   5280
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtproducto 
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   6120
      TabIndex        =   13
      Top             =   4440
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo dc1 
      Height          =   315
      Left            =   7200
      TabIndex        =   11
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
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
      Height          =   315
      Left            =   2040
      TabIndex        =   9
      Top             =   3330
      Width           =   1215
   End
   Begin VB.TextBox txttamaño 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtstock 
      Height          =   405
      Left            =   7200
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtcolor 
      Height          =   405
      Left            =   7200
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton guardar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Guardar producto"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      MaskColor       =   &H00FFC0FF&
      TabIndex        =   5
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
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
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Tamaño"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
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
      Caption         =   "Color:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
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
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nuevo producto"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   2895
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
        lbltipo.Caption = !id_tipoproducto
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
    If txttamaño.Text = "" Then MsgBox "Ingrese el TAMAÑO del producto", vbInformation, "Aviso": txttamaño.SetFocus: Exit Sub
    If Val(txtprecio.Text) = 0 Then MsgBox "Ingrese el PRECIO del producto", vbInformation, "Aviso": txtprecio.SetFocus: Exit Sub
    If dc1.Text = "" Then MsgBox "Seleccione el TIPO del producto", vbInformation, "Aviso": dc1.SetFocus: Exit Sub
    With RsProductos
        .Requery
        .AddNew
        !Producto = txtproducto.Text
        !Color = Txtcolor.Text
        !Stock = Txtstock.Text
        !Precio = txtprecio.Text
        !Tamaño = txttamaño.Text
        !id_tipoproducto = lbltipo.Caption
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
    txttamaño.Text = ""
    dc1.Text = ""
End Sub

