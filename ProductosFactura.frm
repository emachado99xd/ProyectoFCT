VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ProductosFactura 
   Caption         =   "Form2"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11010
   LinkTopic       =   "Form2"
   ScaleHeight     =   7935
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txttotal 
      Height          =   375
      Left            =   7680
      TabIndex        =   18
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox txtdescripcion 
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdagregar 
      Caption         =   "Agregar"
      Height          =   615
      Left            =   4440
      TabIndex        =   14
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox txtcantidad 
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox txtprecio 
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox txtp 
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   4560
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
      Left            =   4920
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txttamaño 
      Height          =   375
      Left            =   8760
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtid 
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   1440
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
      Left            =   9600
      TabIndex        =   19
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      TabIndex        =   17
      Top             =   5280
      Width           =   675
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   5880
      Width           =   1485
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   12
      Top             =   5880
      Width           =   1140
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Producto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   4560
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Precio:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   8
      Top             =   4560
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Id:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Producto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tamaño:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7680
      TabIndex        =   2
      Top             =   840
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Productos para la Factura"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   3930
   End
End
Attribute VB_Name = "ProductosFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdagregar_Click()
    With RsDetalleFactura
        .Requery
        .AddNew
        !Id_Factura = Val(lblfactura.Caption)
        !Id_Producto = Val(lblidproducto.Caption)
        !Cantidad = Val(txtcantidad.Text)
        !Precio = Val(txtprecio.Text)
        .UpdateBatch
    End With
End Sub

Private Sub DataGrid1_Click()
    txtp.Text = DataGrid1.Columns(2).Text
    txtdescripcion.Text = DataGrid1.Columns(6).Text
    txtprecio.Text = DataGrid1.Columns(7).Text
    lblidproducto.Caption = DataGrid1.Columns(0).Text
    End Sub

Private Sub Form_Load()
    Producto
    DetalleFactura
    Adodc1.CursorLocation = adUseClient
    Adodc1.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source= " & App.Path & "\Base\Base.accdb;Persist Security Info=False"
    Adodc1.RecordSource = "SELECT * FROM Cliente"
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
    With Form1
        lblfactura.Caption = .lblfactura.Caption
    End With
End Sub

Private Sub txtcantidad_Change()
    txttotal.Text = Val(txtprecio.Text) * Val(txtcantidad.Text)
End Sub

Private Sub txtid_Change()
    Dim buscar As String
    buscar = "%" & txtid.Text & "%"
    Adodc1.RecordSource = "SELECT *FROM Producto Where [Id_Producto]Like '" & buscar & "'"
    Adodc1.Refresh
End Sub

Private Sub txtproducto_Change()
    Dim buscar As String
    buscar = "%" & txtproducto.Text & "%"
    Adodc1.RecordSource = "SELECT *FROM Producto Where [Producto]Like '" & buscar & "'"
    Adodc1.Refresh
End Sub

Private Sub txttamaño_Change()
    Dim buscar As String
    buscar = "%" & txttamaño.Text & "%"
    Adodc1.RecordSource = "SELECT *FROM Producto Where [Tamaño]Like '" & buscar & "'"
    Adodc1.Refresh
End Sub
