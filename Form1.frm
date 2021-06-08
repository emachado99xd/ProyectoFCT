VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdguardar 
      Caption         =   "Guardar Factura"
      Height          =   615
      Left            =   7200
      TabIndex        =   22
      Top             =   7320
      Width           =   2415
   End
   Begin VB.TextBox txttotal 
      Height          =   375
      Left            =   5280
      TabIndex        =   21
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox txtiva 
      Height          =   405
      Left            =   1080
      TabIndex        =   20
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox txtsubtotal 
      Height          =   375
      Left            =   1800
      TabIndex        =   19
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "++"
      Height          =   615
      Left            =   240
      TabIndex        =   13
      Top             =   2880
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3135
      Left            =   240
      TabIndex        =   11
      Top             =   3720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5530
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
   Begin VB.TextBox txtid 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   11520
      Top             =   4320
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      Connect         =   $"Form1.frx":0000
      OLEDBString     =   $"Form1.frx":0098
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
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
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
      TabIndex        =   18
      Top             =   7320
      Width           =   1365
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Iva:"
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
      TabIndex        =   17
      Top             =   7920
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
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
      Left            =   4080
      TabIndex        =   16
      Top             =   7320
      Width           =   930
   End
   Begin VB.Label lblid 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   12480
      TabIndex        =   15
      Top             =   720
      Width           =   105
   End
   Begin VB.Label lblfactura 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   12720
      TabIndex        =   14
      Top             =   1320
      Width           =   90
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Productos"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4560
      TabIndex        =   12
      Top             =   3120
      Width           =   2010
   End
   Begin VB.Label lbltelefono 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5040
      TabIndex        =   10
      Top             =   1680
      Width           =   90
   End
   Begin VB.Label lbldireccion 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8040
      TabIndex        =   9
      Top             =   1800
      Width           =   90
   End
   Begin VB.Label lblapellido 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5400
      TabIndex        =   8
      Top             =   2520
      Width           =   90
   End
   Begin VB.Label lblnombre 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      TabIndex        =   7
      Top             =   2400
      Width           =   90
   End
   Begin VB.Label Label7 
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
      Left            =   3720
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label6 
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
      Left            =   6360
      TabIndex        =   4
      Top             =   1680
      Width           =   1590
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cedula:"
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
      TabIndex        =   3
      Top             =   1680
      Width           =   1155
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
      Left            =   3600
      TabIndex        =   2
      Top             =   1680
      Width           =   1365
   End
   Begin VB.Label Label3 
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
      TabIndex        =   1
      Top             =   2280
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "factura"
      BeginProperty Font 
         Name            =   "Moonbright Demo"
         Size            =   62.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   3600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdguardar_Click()
Dim Productos As Integer
    With RsFactura
        .Requery
        .Find "Id_Factura='" & Trim(lblfactura.Caption) & "'"
        !Fecha = Date
        !Hora = Time
        !Subtotal = Val(txtsubtotal.Text)
        !Iva = Val(txtiva.Text)
        !Total = Val(txttotal.Text)
        .UpdateBatch
    End With
    Productos = RsTemporal.RecordCount
    RsTemporal.Requery
    RsTemporal.MoveFirst
    For x = 1 To Productos
        With RsDetalleFactura
            .Requery
            .AddNew
            !Id_Factura = DataGrid1.Columns(1).Text
            !Id_Producto = DataGrid1.Columns(2).Text
            !Producto = DataGrid1.Columns(3).Text
            !Cantidad = DataGrid1.Columns(4).Text
            !Precio = DataGrid1.Columns(5).Text
            !Total = DataGrid1.Columns(6).Text
            .UpdateBatch
        End With
        If x = Productos Then Else RsTemporal.MoveNext
    Next
    MsgBox "La factura ha sido creada", vbInformation, "Aviso"
    Subtotal = 0
    borrar
    limpiar
    DataGrid
End Sub

Private Sub Command1_Click()
Productos.Show
End Sub

Private Sub Command2_Click()
    Cliente.Show
End Sub

Private Sub Command3_Click()
    If lblnombre.Caption = "" Then MsgBox "Ingrese un cliente", vbInformation, "Aviso": Txtid.SetFocus: Exit Sub
    If RsTemporal.State = 1 Then RsTemporal.Close
    ProductosFactura.Show
End Sub

Private Sub Command4_Click()
    ModificarCliente.Show
End Sub

Private Sub Command5_Click()
    ModificarProducto.Show
End Sub



Private Sub Form_Load()
Clientes
Factura
Producto
Temporal
DetalleFactura
Set DataGrid1.DataSource = RsTemporal
borrar
DataGrid
End Sub
Private Sub txtid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Txtid.Text = "" Then Exit Sub
        With RsCliente
            .Requery
            .Find "Cedula='" & Trim(Txtid.Text) & "'"
            If .BOF Or .EOF Then Exit Sub
            lblnombre.Caption = !Nombre
            lblapellido.Caption = !Apellido
            lblid.Caption = !Id_Cliente
            lbldireccion.Caption = !Direccion
            lbltelefono.Caption = !Telefono
        End With
        With RsFactura
            .Requery
            .AddNew
            !Id_Cliente = Val(lblid.Caption)
            .UpdateBatch
            lblfactura.Caption = !Id_Factura
        End With
    End If
End Sub

Sub borrar()
    With RsTemporal
        .Requery
        If .BOF Or .EOF Then Exit Sub
        For x = 1 To .RecordCount
            .Delete
            .UpdateBatch
            If .BOF Or .EOF Then Exit Sub
            .MoveNext
        Next
    End With
End Sub
Sub DataGrid()
    DataGrid1.Columns(0).Width = 0
    DataGrid1.Columns(1).Width = 0
    DataGrid1.Columns(2).Width = 0
    DataGrid1.Columns(3).Width = 5150
    DataGrid1.Columns(4).Width = 1500
    DataGrid1.Columns(5).Width = 1500
    DataGrid1.Columns(6).Width = 1500
End Sub
Sub limpiar()
    Txtid.Text = ""
    lblnombre.Caption = ""
    lblapellido.Caption = ""
    lbltelefono.Caption = ""
    lbldireccion.Caption = ""
    lblfactura.Caption = ""
    lblid.Caption = ""
    txtsubtotal.Text = ""
    txtiva.Text = ""
    txttotal.Text = ""
End Sub

