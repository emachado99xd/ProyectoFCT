VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ProductosFactura 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form2"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10740
   LinkTopic       =   "Form2"
   ScaleHeight     =   6900
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   21
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txttotal 
      Height          =   375
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox txtdescripcion 
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdagregar 
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   14
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox txtcantidad 
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txtprecio 
      Height          =   375
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox txtp 
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4800
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   12000
      Top             =   1320
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"ProductosFactura.frx":0000
      OLEDBString     =   $"ProductosFactura.frx":0098
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtproducto 
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txttamaño 
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtid 
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   4895
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblfactura 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   0
      Width           =   45
   End
   Begin VB.Label lblidproducto 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   12240
      TabIndex        =   19
      Top             =   600
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6720
      TabIndex        =   17
      Top             =   4800
      Width           =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   16
      Top             =   5640
      Width           =   1710
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   12
      Top             =   5640
      Width           =   1320
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3720
      TabIndex        =   8
      Top             =   4800
      Width           =   930
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Id:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   3
      Top             =   1080
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tamaño:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6840
      TabIndex        =   2
      Top             =   1080
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "productos para la factura"
      BeginProperty Font 
         Name            =   "Garden Grown US B"
         Size            =   33.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   5430
   End
End
Attribute VB_Name = "ProductosFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Subtotal As Integer
Dim Total, Iva As Double
Private Sub cmdagregar_Click()
    If Txtp.Text = "" Then MsgBox "Seleccione un producto", vbInformation, "Aviso": Exit Sub
    If txtcantidad.Text = "" Then MsgBox "Indique la cantidad que desea comprar", vbInformation, "Aviso": txtcantidad.SetFocus: Exit Sub
     With RsTemporal
        .Requery
        .Find "Id_Producto='" & Trim(lblIdProducto.Caption) & "'"
        If .EOF Then
            .AddNew
            !Id_Factura = Val(lblfactura.Caption)
            !Id_Producto = Val(lblIdProducto.Caption)
            !Producto = Txtp.Text
            !Cantidad = Val(txtcantidad.Text)
            !Precio = Val(Txtprecio.Text)
        Else
            If Val(!Id_Producto) = Val(lblIdProducto.Caption) Then
                !Cantidad = !Cantidad + Val(txtcantidad.Text)
            End If
        End If
        .UpdateBatch
        txttotal.Text = Val(txttotal.Text) + Val(Txtprecio.Text) * Val(txtcantidad.Text)
        Subtotal = Val(txttotal.Text)
        Iva = Subtotal * 0.12
        Total = Subtotal + Iva
    End With
    limpiar
    DataGrid
End Sub
Private Sub Command1_Click()
    With Form1
        RsTemporal.Requery
        Set .DataGrid1.DataSource = RsTemporal
        .txtsubtotal = Subtotal
        .txtiva = Iva
        .txttotal = Total
    End With
    Unload Me
End Sub

Private Sub DataGrid1_Click()
    Txtp.Text = DataGrid1.Columns(2).Text
    Txtdescripcion.Text = DataGrid1.Columns(6).Text
    Txtprecio.Text = DataGrid1.Columns(7).Text
    lblIdProducto.Caption = DataGrid1.Columns(0).Text
    End Sub

Private Sub Form_Load()
    Producto
    Temporal
    Adodc1.CursorLocation = adUseClient
    Adodc1.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source= " & App.Path & "\Base\Base.accdb;Persist Security Info=False"
    Adodc1.RecordSource = "SELECT * FROM Producto"
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
    With Form1
        lblfactura.Caption = .lblfactura.Caption
        If Val(.txtsubtotal.Text) = 0 Then
            Subtotal = 0
        End If
    End With
    txttotal.Text = Subtotal
    DataGrid
End Sub


Private Sub Form_Unload(Cancel As Integer)
    With Form1
        RsTemporal.Requery
        Set .DataGrid1.DataSource = RsTemporal
        .txtsubtotal = Subtotal
        .txtiva = Iva
        .txttotal = Total
    End With
End Sub

Private Sub txtid_Change()
    Dim buscar As String
    buscar = "%" & Txtid.Text & "%"
    Adodc1.RecordSource = "SELECT *FROM Producto Where [Id_Producto]Like '" & buscar & "'"
    Adodc1.Refresh
    DataGrid
End Sub

Private Sub txtproducto_Change()
    Dim buscar As String
    buscar = "%" & Txtproducto.Text & "%"
    Adodc1.RecordSource = "SELECT *FROM Producto Where [Producto]Like '" & buscar & "'"
    Adodc1.Refresh
    DataGrid
End Sub

Private Sub txttamaño_Change()
    Dim buscar As String
    buscar = "%" & Txttamaño.Text & "%"
    Adodc1.RecordSource = "SELECT *FROM Producto Where [Tamaño]Like '" & buscar & "'"
    Adodc1.Refresh
    DataGrid
End Sub
Sub limpiar()
    Txtp.Text = ""
    Txtdescripcion.Text = ""
    Txtprecio.Text = ""
    txtcantidad.Text = ""
End Sub
Sub DataGrid()
    DataGrid1.Columns(0).Width = 0
    DataGrid1.Columns(1).Width = 0
    DataGrid1.Columns(2).Width = 2620
    DataGrid1.Columns(3).Width = 600
    DataGrid1.Columns(4).Width = 1800
    DataGrid1.Columns(5).Width = 1800
    DataGrid1.Columns(6).Width = 2600
    DataGrid1.Columns(7).Width = 600
End Sub
