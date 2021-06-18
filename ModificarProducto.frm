VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form ModificarProducto 
   BackColor       =   &H80000016&
   Caption         =   "Form2"
   ClientHeight    =   9255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9945
   BeginProperty Font 
      Name            =   "MV Boli"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   9255
   ScaleWidth      =   9945
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
      Height          =   375
      Left            =   7440
      TabIndex        =   27
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdeliminar 
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8280
      TabIndex        =   26
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdmodificar 
      Caption         =   "Modificar Porducto"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8280
      TabIndex        =   25
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdnuevo 
      Caption         =   "Nuevo Producto"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8280
      TabIndex        =   24
      Top             =   840
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar Productos"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   975
      Left            =   240
      TabIndex        =   19
      Top             =   720
      Width           =   7935
      Begin VB.TextBox txttipo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Unispace"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   29
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtproducto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Unispace"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   23
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtid 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Unispace"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   480
         TabIndex        =   22
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   300
         Left            =   5400
         TabIndex        =   28
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   300
         Left            =   1680
         TabIndex        =   21
         Top             =   360
         Width           =   1410
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Id:"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   300
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdguardar 
      BackColor       =   &H80000005&
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8280
      TabIndex        =   18
      Top             =   2280
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   16560
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
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
      Connect         =   $"ModificarProducto.frx":0000
      OLEDBString     =   $"ModificarProducto.frx":0094
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   720
      Left            =   1680
      TabIndex        =   16
      Top             =   8040
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1270
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Showcard Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Txtprecio 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2280
      TabIndex        =   13
      Top             =   7320
      Width           =   5895
   End
   Begin VB.TextBox Txtdescripcion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   12
      Top             =   6600
      Width           =   4335
   End
   Begin VB.TextBox Txtcolor 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2160
      TabIndex        =   11
      Top             =   5880
      Width           =   6015
   End
   Begin VB.TextBox Txtstock 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   10
      Top             =   5160
      Width           =   6015
   End
   Begin VB.TextBox Txttamaño 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   9
      Top             =   4440
      Width           =   5535
   End
   Begin VB.TextBox Txtp 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3360
      TabIndex        =   8
      Top             =   3720
      Width           =   4815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   3201
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
            LCID            =   12298
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
            LCID            =   12298
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
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
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
      Left            =   240
      TabIndex        =   17
      Top             =   8040
      Width           =   1425
   End
   Begin VB.Label lblIdTipo 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15960
      TabIndex        =   15
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lblIdProducto 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16440
      TabIndex        =   14
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
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
      Left            =   240
      TabIndex        =   7
      Top             =   6600
      Width           =   3540
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Precio:"
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
      Left            =   240
      TabIndex        =   6
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tamaño:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   4440
      Width           =   2310
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Color:"
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
      Left            =   240
      TabIndex        =   4
      Top             =   5880
      Width           =   1890
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Stock:"
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
      Left            =   240
      TabIndex        =   3
      Top             =   5160
      Width           =   1860
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   3030
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Control de productos"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   660
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   6660
   End
End
Attribute VB_Name = "ModificarProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdeliminar_Click()
    If txtp.Text = "" Then MsgBox "Seleccione un producto", vbInformation, "Aviso": Exit Sub
    If cmdeliminar.Caption = "Eliminar" Then
        With RsProductos
            .Requery
            .Find "Id_Producto='" & Trim(lblidproducto.Caption) & "'"
            If .EOF Then Exit Sub
            .Delete
            .UpdateBatch
            .Requery
        End With
        limpiar
        Adodc1.Refresh
        DataGrid
        bloquear True
        txtid.Locked = False
        txtproducto.Locked = False
    Else
        txtid.Locked = False
        txtproducto.Locked = False
        cmdnuevo.Enabled = True
        cmdmodificar.Enabled = True
        cmdguardar.Enabled = False
        cmdeliminar.Caption = "Eliminar"
        bloquear True
        limpiar
    End If
    DataGrid
End Sub

Private Sub cmdguardar_Click()
    With RsProductos
        .Requery
        .Find "Id_producto ='" & Trim(lblidproducto.Caption) & "'"
        !Producto = txtp.Text
        !Tipo = DataCombo1.Text
        !Stock = Val(Txtstock.Text)
        !Color = Txtcolor.Text
        !Tamaño = Txttamaño.Text
        !Descripcion = txtdescripcion.Text
        !Precio = Val(txtprecio.Text)
        With RsTipoProducto
            .Requery
            .Find "Tipo='" & Trim(DataCombo1.Text) & "'"
            lblIdTipo.Caption = !Id_Tipoproducto
        End With
            !Id_Tipoproducto = lblIdTipo.Caption
            .UpdateBatch
    End With
        MsgBox "El producto ha sido modificado correctamente", vbInformation, aviso
        bloquear True
        txtid.Locked = False
        txtproducto.Locked = False
        cmdnuevo.Enabled = True
        cmdmodificar.Enabled = True
        cmdguardar.Enabled = False
        cmdeliminar.Caption = "Eliminar"
        limpiar
        Adodc1.Refresh
        DataGrid
End Sub

Private Sub cmdmodificar_Click()
    If txtp.Text = "" Then MsgBox "Seleccione un producto", vbInformation, "Aviso": Exit Sub
    bloquear False
    txtid.Locked = True
    txtproducto.Locked = True
    cmdnuevo.Enabled = False
    cmdmodificar.Enabled = False
    cmdguardar.Enabled = True
    cmdeliminar.Caption = "Cancelar"
    DataGrid
End Sub

Private Sub cmdnuevo_Click()
    Productos.Show
    limpiar
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub DataGrid1_Click()
    txtp.Text = DataGrid1.Columns(2).Text
    Txtstock.Text = DataGrid1.Columns(4).Text
    Txtcolor.Text = DataGrid1.Columns(5).Text
    Txttamaño.Text = DataGrid1.Columns(6).Text
    txtdescripcion.Text = DataGrid1.Columns(7).Text
    txtprecio.Text = DataGrid1.Columns(8).Text
    lblidproducto.Caption = DataGrid1.Columns(0).Text
    lblIdTipo.Caption = DataGrid1.Columns(1).Text
    With RsTipoProducto
        .Requery
        .Find "Id_tipoproducto='" & Trim(lblIdTipo.Caption) & "'"
        DataCombo1.Text = !Tipo
    End With
    txtid.Locked = False
    txtproducto.Locked = False
    cmdnuevo.Enabled = True
    cmdmodificar.Enabled = True
    cmdguardar.Enabled = False
    cmdeliminar.Caption = "Eliminar"
    bloquear True
End Sub


Private Sub Form_Activate()
    Adodc1.Refresh
    DataGrid
End Sub

Private Sub Form_Load()
    ModificarProducto.Picture = LoadPicture(App.Path & "\IMG\Fondo1.jpg")
    Producto
    TipoProducto
    Adodc1.CursorLocation = adUseClient
    Adodc1.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source= " & App.Path & "\Base\Base.accdb;Persist Security Info=False"
    Adodc1.RecordSource = "SELECT * FROM Producto"
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
    Set DataCombo1.RowSource = RsTipoProducto
    DataCombo1.BoundColumn = "Tipo"
    DataCombo1.ListField = "Tipo"
    DataGrid
    bloquear True
    cmdguardar.Enabled = False
End Sub

Private Sub txtid_Change()
    Dim buscar As String
    buscar = "%" & txtid.Text & "%"
    Adodc1.RecordSource = "SELECT *FROM Producto Where [Id_Producto]Like '" & buscar & "'"
    Adodc1.Refresh
    DataGrid
End Sub

Private Sub txtproducto_Change()
    Dim buscar As String
    buscar = "%" & txtproducto.Text & "%"
    Adodc1.RecordSource = "SELECT *FROM Producto Where [Producto]Like '" & buscar & "'"
    Adodc1.Refresh
    DataGrid
End Sub

Sub DataGrid()
    DataGrid1.Columns(0).Width = 0
    DataGrid1.Columns(1).Width = 0
    DataGrid1.Columns(2).Width = 1250
    DataGrid1.Columns(3).Width = 1240
    DataGrid1.Columns(4).Width = 1200
    DataGrid1.Columns(5).Width = 1000
    DataGrid1.Columns(6).Width = 1000
    DataGrid1.Columns(7).Width = 1000
    DataGrid1.Columns(8).Width = 900
End Sub

Sub bloquear(estado As Boolean)
    txtp.Locked = estado
    Txttamaño.Locked = estado
    Txtstock.Locked = estado
    Txtcolor.Locked = estado
    txtdescripcion.Locked = estado
    txtprecio.Locked = estado
    DataCombo1.Locked = estado
End Sub

Sub limpiar()
    txtproducto.Text = ""
    txtid.Text = ""
    txtp.Text = ""
    Txttamaño.Text = ""
    Txtstock.Text = ""
    Txtcolor.Text = ""
    txtdescripcion.Text = ""
    txtprecio.Text = ""
    DataCombo1.Text = ""
End Sub

Private Sub txttipo_Change()
    Dim buscar As String
    buscar = "%" & txttipo.Text & "%"
    Adodc1.RecordSource = "SELECT *FROM Producto Where [Tipo]Like '" & buscar & "'"
    Adodc1.Refresh
    DataGrid
End Sub
