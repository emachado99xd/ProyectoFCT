VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Anular 
   Caption         =   "Form2"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7365
   LinkTopic       =   "Form2"
   ScaleHeight     =   10935
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Mostrar Reporte"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   14
      Top             =   9840
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Reportes"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   12
      Top             =   7800
      Width           =   5295
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text3 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "d/M/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   22538
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2760
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "d/M/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   22538
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "Anular.frx":0000
         Left            =   1200
         List            =   "Anular.frx":000D
         TabIndex        =   13
         Text            =   "Opciones"
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1920
         TabIndex        =   21
         Top             =   1680
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3120
         TabIndex        =   19
         Top             =   1560
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   840
         TabIndex        =   18
         Top             =   1560
         Visible         =   0   'False
         Width           =   900
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Validar"
      Height          =   735
      Left            =   2880
      TabIndex        =   10
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Anular"
      Height          =   735
      Left            =   1200
      TabIndex        =   8
      Top             =   6840
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
      Connect         =   $"Anular.frx":002A
      OLEDBString     =   $"Anular.frx":00C2
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
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Width           =   5535
      _ExtentX        =   9763
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
      Connect         =   $"Anular.frx":015A
      OLEDBString     =   $"Anular.frx":01F2
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
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   5535
      _ExtentX        =   9763
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
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1095
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   3855
      Begin VB.TextBox txtcedula 
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Cédula:"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2160
      TabIndex        =   20
      Top             =   9360
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Productos de la factura"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   360
      TabIndex        =   11
      Top             =   4440
      Width           =   4815
   End
   Begin VB.Label l 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   4800
      TabIndex        =   9
      Top             =   1320
      Width           =   60
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
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Facturas"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   660
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   2595
   End
End
Attribute VB_Name = "Anular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IdProducto, Cantidad, f As Integer
Dim fecha, Inicio As Date



Private Sub Combo1_Click()
    If Combo1.Text = "Fecha" Then
        Text2.Visible = True
        Text3.Visible = True
        Text4.Visible = False
        Combo1.Top = 600
        Command3.Enabled = True
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Label4.Visible = True
        Label5.Visible = True
        Label7.Visible = False
    Else
        If Combo1.Text = "Anuladas" Then
            Text2.Visible = True
            Text3.Visible = True
            Text4.Visible = False
            Combo1.Top = 600
            Command3.Enabled = True
            Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            Label4.Visible = True
            Label5.Visible = True
            Label7.Visible = False
        Else
            Text2.Visible = False
            Text3.Visible = False
            Text4.Visible = True
            Combo1.Top = 600
            Command3.Enabled = True
            Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            Label4.Visible = False
            Label5.Visible = False
            Label7.Visible = True
        End If
    End If
End Sub

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
        For X = 1 To .RecordCount
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
    With RsFactura
        .Requery
        .Find "Id_Factura='" & Trim(f) & "'"
        If .BOF Or .EOF Then Exit Sub
        If !Validar = True Then
            l.Caption = "Factura Valida"
        Else
            l.Caption = "Factura No válida"
        End If
        .Requery
    End With
    MsgBox "La factura a sido anulada", vbInformation, "Aviso"
    DataGrid
    DataGrida
End Sub

Private Sub Command2_Click()
    With RsFactura
        .Requery
        .Find "Id_Factura='" & Trim(lblfactura.Caption) & "'"
        If !Validar = True Then MsgBox "Esta factura ya es válida", vbInformation, "Aviso": Exit Sub
        !Validar = True
        .UpdateBatch
    End With
    Dim buscar As String
    buscar = "%" & lblfactura.Caption & "%"
    Adodc2.RecordSource = "SELECT *FROM Detalle_Factura Where [Id_Factura]Like '" & buscar & "'"
    Adodc2.Refresh
    Set RsDetalleFactura.DataSource = Adodc2
    With RsDetalleFactura
        For X = 1 To .RecordCount
            IdProducto = !Id_Producto
            Cantidad = !Cantidad
            With RsProductos
                .Requery
                .Find "Id_Producto='" & Trim(IdProducto) & "'"
                !Stock = Val(!Stock) - Cantidad
                .UpdateBatch
            End With
            .MoveNext
            If .BOF Then Exit Sub
        Next
    End With
    With RsFactura
        .Requery
        .Find "Id_Factura='" & Trim(f) & "'"
        If .BOF Or .EOF Then Exit Sub
        If !Validar = True Then
            l.Caption = "Factura Valida"
        Else
            l.Caption = "Factura No válida"
        End If
        .Requery
    End With
    MsgBox "La factura a sido validada", vbInformation, "Aviso"
    DataGrid
    DataGrida
End Sub

Private Sub Command3_Click()
    If Combo1.Text = "Fecha" Then
    If Text2.Text = "" Or Text3.Text = "" Then Exit Sub
    Inicio = Text2.Text
    fecha = Text3.Text
    a = "#" & Inicio & "#"
    b = "#" & fecha & "#"
        With RsFactura
           ' If .State = 1 Then .Close
           ' a = "#" & Inicio & "#"
           ' .Open "Select * From Factura Where ((Factura.[Fecha])= " & a & ")", Base, adOpenStatic, adLockBatchOptimistic
           ' If .EOF Or .BOF Then MsgBox "Fecha(a) no hay": Exit Sub
            'b = "#" & fecha & "#"
           ' If .State = 1 Then .Close
             '   .Open "Select * From Factura Where ((Factura.[Fecha])= " & b & ")", Base, adOpenStatic, adLockBatchOptimistic
           ' If .EOF Or .BOF Then MsgBox "Fecha(b) no hay": Exit Sub
            If .State = 1 Then .Close
                .Open "Select * From Factura Where ((Factura.[Fecha])>= " & a & ") AND ((Factura.[Fecha])<= " & b & ") AND ((Factura.[Validar])= TRUE) ", Base, adOpenStatic, adLockBatchOptimistic
            Set DataReport2.DataSource = RsFactura
            DataReport2.Show
        End With
    Else
        If Combo1.Text = "Anuladas" Then
        If Text2.Text = "" Or Text3.Text = "" Then Exit Sub
        Inicio = Text2.Text
    fecha = Text3.Text
    a = "#" & Inicio & "#"
    b = "#" & fecha & "#"
            With RsFactura
                If .State = 1 Then .Close
               
                .Open "Select * From Factura Where ((Factura.[Fecha])>= " & a & ") AND ((Factura.[Fecha])<= " & b & ") AND ((Factura.[Validar])= False) ", Base, adOpenStatic, adLockBatchOptimistic
                
                '.Open "Select * From Factura Where ((Factura.[Validar])= False)", Base, adOpenStatic, adLockBatchOptimistic
                Set DataReport2.DataSource = RsFactura
                DataReport2.Show
            End With
        Else
            If Combo1.Text = "Precio" Then
                With RsFactura
                    o = Val(Text4.Text)
                    If .State = 1 Then .Close
                    '.Open "Select * From Factura Where ((Factura.[Fecha])>= " & a & ") AND ((Factura.[Fecha])<= " & b & ") AND (Factura.[Total])> " & o & ")", Base, adOpenStatic, adLockBatchOptimistic
                    .Open "Select * From Factura Where ((Factura.[Total])> " & o & ") AND ((Factura.[Validar])= True)", Base, adOpenStatic, adLockBatchOptimistic
                    Set DataReport2.DataSource = RsFactura
                    DataReport2.Show
                End With
            End If
        End If
    End If
End Sub

Private Sub DataGrid1_Click()
    lblfactura.Caption = DataGrid1.Columns(0).Text
    Dim buscar As String
    buscar = "%" & lblfactura.Caption & "%"
    Adodc2.RecordSource = "SELECT *FROM Detalle_Factura Where [Id_Factura]Like '" & buscar & "'"
    Adodc2.Refresh
    Set DataGrid2.DataSource = Adodc2
    f = DataGrid1.Columns(0).Text
    With RsFactura
        .Requery
        .Find "Id_Factura='" & Trim(f) & "'"
        If .BOF Or .EOF Then Exit Sub
        If !Validar = True Then
            l.Caption = "Factura Valida"
        Else
            l.Caption = "Factura No válida"
        End If
        .Requery
    End With
    DataGrid2.Enabled = True
    DataGrid
    DataGrida
End Sub



Private Sub Form_Load()
    Anular.Picture = LoadPicture(App.Path & "\IMG\Fondo1.jpg")
    Clientes
    Factura
    DetalleFactura
    Producto
    Adodc1.CursorLocation = adUseClient
    Adodc1.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source= " & App.Path & "\Base\Base.accdb;Persist Security Info=False"
    Adodc1.RecordSource = "SELECT * FROM Factura"
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
    Adodc2.CursorLocation = adUseClient
    Adodc2.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source= " & App.Path & "\Base\Base.accdb;Persist Security Info=False"
    Adodc2.RecordSource = "SELECT * FROM Detalle_Factura"
    Adodc2.Refresh
    DataGrid1.Enabled = False
    DataGrid2.Enabled = False
    DataGrid
End Sub


Private Sub txtcedula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With RsCliente
            .Requery
            .Find "Cedula='" & Trim(txtcedula.Text) & "'"
            If .BOF Or .EOF Then Exit Sub
            lblid.Caption = !Id_Cliente
        End With
    Dim buscar As String
    buscar = "%" & lblid.Caption & "%"
    Adodc1.RecordSource = "SELECT *FROM Factura Where [Id_Cliente]Like '" & buscar & "'"
    Adodc1.Refresh
    DataGrid1.Enabled = True
    DataGrid
    End If
End Sub

Sub DataGrid()
    DataGrid1.Columns(0).Width = 0
    DataGrid1.Columns(1).Width = 0
    DataGrid1.Columns(2).Width = 1000
    DataGrid1.Columns(3).Width = 1300
    DataGrid1.Columns(4).Width = 800
    DataGrid1.Columns(5).Width = 800
    DataGrid1.Columns(6).Width = 1068
    DataGrid1.Columns(7).Width = 0
End Sub
Sub DataGrida()
    DataGrid2.Columns(0).Width = 0
    DataGrid2.Columns(1).Width = 0
    DataGrid2.Columns(2).Width = 0
    DataGrid2.Columns(3).Width = 2200
    DataGrid2.Columns(4).Width = 994
    DataGrid2.Columns(5).Width = 994
    DataGrid2.Columns(6).Width = 994
End Sub
