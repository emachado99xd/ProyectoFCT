VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form ModificarProducto 
   BackColor       =   &H80000016&
   Caption         =   "Form2"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11100
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
   ScaleHeight     =   7560
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdguardar 
      BackColor       =   &H80000005&
      Caption         =   "guardar"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   23
      Top             =   6840
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9360
      Top             =   240
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
      Height          =   315
      Left            =   3360
      TabIndex        =   20
      Top             =   6480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Txtproducto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   17
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Txtid 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1800
      TabIndex        =   16
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Txtprecio 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6720
      TabIndex        =   14
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox Txtdescripcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Txtcolor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6720
      TabIndex        =   12
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Txtstock 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      TabIndex        =   11
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox Txttamaño 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Txtp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      TabIndex        =   9
      Top             =   3600
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1815
      Left            =   480
      TabIndex        =   2
      Top             =   1560
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
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   15
      Left            =   9480
      TabIndex        =   24
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000005&
      Caption         =   "tipo"
      Height          =   375
      Left            =   2160
      TabIndex        =   22
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   21
      Top             =   1560
      Width           =   2295
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
      Left            =   8400
      TabIndex        =   19
      Top             =   720
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
      Left            =   9840
      TabIndex        =   18
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      Caption         =   "Prodcuto"
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000005&
      Caption         =   "Descripción"
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000005&
      Caption         =   "Precio"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000005&
      Caption         =   "tamaño"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      Caption         =   "color"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      Caption         =   "stock"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000005&
      Caption         =   "Producto"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "ID"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Modificar producto"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "ModificarProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdguardar_Click()
    With RsProductos
        .Requery
        .Find "Id_producto ='" & Trim(lblIdTipo.Caption) & "'"
        !Producto = Txtp.Text
        !Stock = Val(Txtstock.Text)
        !Color = Txtcolor.Text
        !Tamaño = Txttamaño.Text
        !Descripción = Txtdescripcion.Text
        !Precio = Val(Txtprecio.Text)
        With RsTipoProducto
            .Requery
            .Find "tipo ='" & Trim(DataCombo1.Text) & "'"
            lblIdTipo.Caption = !id_tipoproducto
        End With
            id_tipoproducto = lblIdTipo.Caption
            .UpdateBatch
    End With
        MsgBox "El producto ha sido modificado correctamente", vbInformation, aviso
        Adodc1.Refresh
End Sub

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub DataGrid1_Click()
    Txtp.Text = DataGrid1.Columns(2).Text
    Txtstock.Text = DataGrid1.Columns(3).Text
    Txtcolor.Text = DataGrid1.Columns(4).Text
    Txttamaño.Text = DataGrid1.Columns(5).Text
    Txtdescripcion.Text = DataGrid1.Columns(6).Text
    Txtprecio.Text = DataGrid1.Columns(7).Text
    lblIdProducto.Caption = DataGrid1.Columns(0).Text
    lblIdTipo.Caption = DataGrid1.Columns(1).Text
    With RsTipoProducto
        .Requery
        .Find "Id_tipoproducto='" & Trim(lblIdTipo.Caption) & "'"
        DataCombo1.Text = !tipo
    End With
End Sub

Private Sub Form_Load()
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
End Sub

Private Sub txtid_Change()
    Dim buscar As String
    buscar = "%" & Txtid.Text & "%"
    Adodc1.RecordSource = "SELECT *FROM Producto Where [Id_Producto]Like '" & buscar & "'"
    Adodc1.Refresh
End Sub

Private Sub txtproducto_Change()
    Dim buscar As String
    buscar = "%" & Txtproducto.Text & "%"
    Adodc1.RecordSource = "SELECT *FROM Producto Where [Producto]Like '" & buscar & "'"
    Adodc1.Refresh
End Sub

