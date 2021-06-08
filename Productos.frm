VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Productos 
   BackColor       =   &H80000003&
   Caption         =   "Form2"
   ClientHeight    =   4230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   LinkTopic       =   "Form2"
   ScaleHeight     =   4230
   ScaleWidth      =   7755
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
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   13
      Top             =   3480
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo dc1 
      Height          =   315
      Left            =   6120
      TabIndex        =   11
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
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
      Left            =   1920
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txttamaño 
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtstock 
      Height          =   405
      Left            =   6240
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtcolor 
      Height          =   405
      Left            =   6360
      TabIndex        =   6
      Top             =   960
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
      Left            =   1800
      MaskColor       =   &H00FFC0FF&
      TabIndex        =   5
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
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
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Tamaño"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Color:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nuevo producto"
      BeginProperty Font 
         Name            =   "Holiday Party"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   855
      Left            =   2520
      TabIndex        =   0
      Top             =   -120
      Width           =   4815
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
    If Txtproducto.Text = "" Then MsgBox "Ingrese el NOMBRE del producto", vbInformation, "Aviso": Txtproducto.SetFocus: Exit Sub
    If Txtcolor.Text = "" Then MsgBox "Ingrese el COLOR del producto", vbInformation, "Aviso": Txtcolor.SetFocus: Exit Sub
    If Val(Txtstock.Text) = 0 Then MsgBox "Ingrese el STOCK del producto", vbInformation, "Aviso": Txtstock.SetFocus: Exit Sub
    If Txttamaño.Text = "" Then MsgBox "Ingrese el TAMAÑO del producto", vbInformation, "Aviso": Txttamaño.SetFocus: Exit Sub
    If Val(Txtprecio.Text) = 0 Then MsgBox "Ingrese el PRECIO del producto", vbInformation, "Aviso": Txtprecio.SetFocus: Exit Sub
    If dc1.Text = "" Then MsgBox "Seleccione el TIPO del producto", vbInformation, "Aviso": dc1.SetFocus: Exit Sub
    With RsProductos
        .Requery
        .AddNew
        !Producto = Txtproducto.Text
        !Color = Txtcolor.Text
        !Stock = Txtstock.Text
        !Precio = Txtprecio.Text
        !Tamaño = Txttamaño.Text
        !id_tipoproducto = lbltipo.Caption
        .UpdateBatch
End With
MsgBox "El producto ha sido registrado correctamente", vbInformation, "Aviso"
limpiar
End Sub

Sub limpiar()
    Txtproducto.Text = ""
    Txtcolor.Text = ""
    Txtstock.Text = ""
    Txtprecio.Text = ""
    Txttamaño.Text = ""
    dc1.Text = ""
End Sub

