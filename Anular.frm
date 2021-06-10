VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Anular 
   Caption         =   "Form2"
   ClientHeight    =   8670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11460
   LinkTopic       =   "Form2"
   ScaleHeight     =   8670
   ScaleWidth      =   11460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Anular"
      Height          =   735
      Left            =   4080
      TabIndex        =   8
      Top             =   6120
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   12480
      Top             =   4080
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
      Connect         =   $"Anular.frx":0000
      OLEDBString     =   $"Anular.frx":0098
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1575
      Left            =   1080
      TabIndex        =   7
      Top             =   4200
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2778
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   12240
      Top             =   1560
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
      Connect         =   $"Anular.frx":0130
      OLEDBString     =   $"Anular.frx":01C8
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2055
      Left            =   1080
      TabIndex        =   4
      Top             =   2040
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   3625
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
   Begin VB.Frame Frame1 
      Caption         =   "Buscar facturas"
      Height          =   1095
      Left            =   3240
      TabIndex        =   0
      Top             =   720
      Width           =   3855
      Begin VB.TextBox txtcedula 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Cédula:"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Label l 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Unispace"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9840
      TabIndex        =   9
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label lblfactura 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   12600
      TabIndex        =   6
      Top             =   3120
      Width           =   45
   End
   Begin VB.Label lblid 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   12600
      TabIndex        =   5
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Facturas"
      Height          =   195
      Left            =   3600
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Anular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IdProducto, Cantidad As Integer
Private Sub Command1_Click()
    With RsFactura
        .Requery
        .Find "Id_Factura='" & Trim(lblfactura.Caption) & "'"
        If !Validar = False Then MsgBox "Esta factura no es válida", vbInformation, "Aviso": Exit Sub
        !Validar = False
        .UpdateBatch
    End With
    Dim buscar As String
    buscar = "%" & lblfactura.Caption & "%"
    Adodc2.RecordSource = "SELECT *FROM Detalle_Factura Where [Id_Factura]Like '" & buscar & "'"
    Adodc2.Refresh
    Set RsDetalleFactura.DataSource = Adodc2
    With RsDetalleFactura
        For x = 1 To .RecordCount
            IdProducto = !Id_Producto
            Cantidad = !Cantidad
            With RsProductos
                .Requery
                .Find "Id_Producto='" & Trim(IdProducto) & "'"
                !Stock = Val(!Stock) + Cantidad
                .UpdateBatch
            End With
            .MoveNext
            If .BOF Then Exit Sub
        Next
    End With
    MsgBox "La factura a sido anulada", vbInformation, "Aviso"
End Sub

Private Sub DataGrid1_Click()
    lblfactura.Caption = DataGrid1.Columns(0).Text
    Dim buscar As String
    buscar = "%" & lblfactura.Caption & "%"
    Adodc2.RecordSource = "SELECT *FROM Detalle_Factura Where [Id_Factura]Like '" & buscar & "'"
    Adodc2.Refresh
    Set DataGrid2.DataSource = Adodc2
    With RsFactura
        If !Validar = True Then
            l.Caption = "Valida"
        Else
            l.Caption = "No válida"
        End If
    End With
End Sub

Private Sub Form_Load()
    Clientes
    Factura
    DetalleFactura
    Adodc1.CursorLocation = adUseClient
    Adodc1.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source= " & App.Path & "\Base\Base.accdb;Persist Security Info=False"
    Adodc1.RecordSource = "SELECT * FROM Factura"
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
    Adodc2.CursorLocation = adUseClient
    Adodc2.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source= " & App.Path & "\Base\Base.accdb;Persist Security Info=False"
    Adodc2.RecordSource = "SELECT * FROM Detalle_Factura"
    Adodc2.Refresh
End Sub


Private Sub txtcedula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With RsCliente
            .Requery
            .Find "Cedula='" & Trim(txtcedula.Text) & "'"
            lblid.Caption = !Id_Cliente
        End With
    End If
    Dim buscar As String
    buscar = "%" & lblid.Caption & "%"
    Adodc1.RecordSource = "SELECT *FROM Factura Where [Id_Cliente]Like '" & buscar & "'"
    Adodc1.Refresh
End Sub
