VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ModificarCliente 
   BackColor       =   &H80000010&
   Caption         =   "Form2"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11430
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   6660
   ScaleWidth      =   11430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   4920
      TabIndex        =   16
      Top             =   5880
      Width           =   1335
   End
   Begin VB.TextBox txttel 
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox txtdir 
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox txtced 
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox txtape 
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox txtnom 
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   4080
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   11640
      Top             =   1080
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
      Connect         =   $"ModificarCliente.frx":0000
      OLEDBString     =   $"ModificarCliente.frx":00A2
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
      Height          =   1815
      Left            =   1800
      TabIndex        =   5
      Top             =   2160
      Width           =   8295
      _ExtentX        =   14631
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
            LCID            =   22538
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
            LCID            =   22538
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
   Begin VB.TextBox txtcedula 
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtnombre 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblid 
      Height          =   255
      Left            =   12000
      TabIndex        =   17
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   10
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cedula"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   8
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   7
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cedula"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "nombre"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Modificar cliente"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
End
Attribute VB_Name = "ModificarCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
With RsCliente
    .Requery
    .Find "Id_Cliente='" & Trim(lblid.Caption) & "'"
    !Nombre = txtnom.Text
    !Apellido = txtape.Text
    !Cedula = txtced.Text
    !Direccion = txtdir.Text
    !Telefono = txttel.Text
    .UpdateBatch
    End With
    MsgBox "El registro ha sido modificado correctamente", vbInformation, "Aviso"
    Adodc1.Refresh
End Sub

Private Sub DataGrid1_Click()
    txtnom.Text = DataGrid1.Columns(1).Text
    txtape.Text = DataGrid1.Columns(2).Text
    txtced.Text = DataGrid1.Columns(3).Text
    txtdir.Text = DataGrid1.Columns(4).Text
    txttel.Text = DataGrid1.Columns(5).Text
    lblid.Caption = DataGrid1.Columns(0).Text
    
End Sub

Private Sub Form_Load()
    Clientes
    Adodc1.CursorLocation = adUseClient
    Adodc1.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source= " & App.Path & "\Base\Base.accdb;Persist Security Info=False"
    Adodc1.RecordSource = "SELECT * FROM Cliente"
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub txtcedula_Change()
    Dim buscar As String
    buscar = "%" & txtcedula.Text & "%"
    Adodc1.RecordSource = "SELECT *FROM Cliente Where [Cedula]Like '" & buscar & "'"
    Adodc1.Refresh
End Sub

Private Sub txtnombre_Change()
    Dim buscar As String
    buscar = "%" & txtnombre.Text & "%"
    Adodc1.RecordSource = "SELECT *FROM Cliente Where [Nombre]Like '" & buscar & "'"
    Adodc1.Refresh
End Sub
