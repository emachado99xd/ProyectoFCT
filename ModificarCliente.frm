VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ModificarCliente 
   BackColor       =   &H8000000B&
   Caption         =   "Form2"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9915
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7950
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      BackColor       =   &H000080FF&
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
      Height          =   735
      Left            =   8160
      TabIndex        =   22
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   21
      Top             =   240
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3201
      _Version        =   393216
      BackColor       =   -2147483644
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
      Caption         =   "Modificar Cliente"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      TabIndex        =   20
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Nuevo Cliente"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      TabIndex        =   19
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar cliente"
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
      Height          =   1215
      Left            =   360
      TabIndex        =   14
      Top             =   720
      Width           =   7815
      Begin VB.TextBox txtcedula 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         TabIndex        =   18
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtnombre 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   15
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cédula:"
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
         Left            =   4080
         TabIndex        =   17
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
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
         Left            =   600
         TabIndex        =   16
         Top             =   600
         Width           =   1110
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
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
      Height          =   735
      Left            =   8160
      TabIndex        =   12
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox txttel 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   11
      Top             =   6240
      Width           =   4815
   End
   Begin VB.TextBox txtdir 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   10
      Top             =   6960
      Width           =   4695
   End
   Begin VB.TextBox txtced 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   9
      Top             =   5520
      Width           =   5175
   End
   Begin VB.TextBox txtape 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   8
      Top             =   4800
      Width           =   4935
   End
   Begin VB.TextBox txtnom 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   7
      Top             =   4080
      Width           =   4935
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
   Begin VB.Label lblid 
      Height          =   255
      Left            =   12000
      TabIndex        =   13
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   6120
      Width           =   2730
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   6840
      Width           =   2925
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cedula:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   2145
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   2310
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Control de clientes"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   5820
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
    bloquear True
    txtnombre.Locked = False
    txtcedula.Locked = False
    limpiar
    DataGrid
    Command1.Enabled = False
    Command2.Enabled = True
    Command3.Enabled = True
    Command5.Caption = "Eliminar"
End Sub

Private Sub Command2_Click()
    Cliente.Show
    limpiar
End Sub

Private Sub Command3_Click()
    If txtnom.Text = "" Then MsgBox "Seleccione un cliente", vbInformation, "Aviso": Exit Sub
    bloquear False
    txtnom.SetFocus
    txtnombre.Locked = True
    txtcedula.Locked = True
    Command1.Enabled = True
    Command2.Enabled = False
    Command3.Enabled = False
    Command5.Caption = "Cancelar"
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
    If txtnom.Text = "" Then MsgBox "Seleccione un cliente", vbInformation, "Aviso": Exit Sub
    If Command5.Caption = "Eliminar" Then
        With RsCliente
            .Requery
            .Find "Id_Cliente='" & Trim(lblid.Caption) & "'"
            If .EOF Then Exit Sub
            .Delete
            .UpdateBatch
            .Requery
        End With
        limpiar
        Adodc1.Refresh
        DataGrid
    Else
        Command1.Enabled = False
        Command2.Enabled = True
        Command3.Enabled = True
        Command5.Caption = "Eliminar"
        bloquear True
        limpiar
    End If
End Sub

Private Sub DataGrid1_Click()
    txtnom.Text = DataGrid1.Columns(1).Text
    txtape.Text = DataGrid1.Columns(2).Text
    txtced.Text = DataGrid1.Columns(3).Text
    txtdir.Text = DataGrid1.Columns(4).Text
    txttel.Text = DataGrid1.Columns(5).Text
    lblid.Caption = DataGrid1.Columns(0).Text
    DataGrid
    Command1.Enabled = False
    Command2.Enabled = True
    Command3.Enabled = True
    Command5.Caption = "Eliminar"
    bloquear True
    txtnombre.Locked = False
    txtcedula.Locked = False
End Sub
Private Sub Form_Activate()
    Adodc1.Refresh
    DataGrid
End Sub
Private Sub Form_Load()
    ModificarCliente.Picture = LoadPicture(App.Path & "\IMG\Fondo1.jpg")
    Clientes
    Adodc1.CursorLocation = adUseClient
    Adodc1.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source= " & App.Path & "\Base\Base.accdb;Persist Security Info=False"
    Adodc1.RecordSource = "SELECT * FROM Cliente"
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
    bloquear True
    limpiar
    DataGrid
    Command1.Enabled = False
End Sub

Private Sub txtcedula_Change()
    Dim buscar As String
    buscar = "%" & txtcedula.Text & "%"
    Adodc1.RecordSource = "SELECT *FROM Cliente Where [Cedula]Like '" & buscar & "'"
    Adodc1.Refresh
    DataGrid
End Sub

Private Sub txtnombre_Change()
    Dim buscar As String
    buscar = "%" & txtnombre.Text & "%"
    Adodc1.RecordSource = "SELECT *FROM Cliente Where [Nombre]Like '" & buscar & "'"
    Adodc1.Refresh
    DataGrid
End Sub

Sub bloquear(estado As Boolean)
    txtnom.Locked = estado
    txtape.Locked = estado
    txtced.Locked = estado
    txttel.Locked = estado
    txtdir.Locked = estado
End Sub
Sub limpiar()
    txtnombre.Text = ""
    txtcedula.Text = ""
    txtnom.Text = ""
    txtape.Text = ""
    txtced.Text = ""
    txttel.Text = ""
    txtdir.Text = ""
End Sub

Sub DataGrid()
    DataGrid1.Columns(0).Width = 0
    DataGrid1.Columns(1).Width = 1600
    DataGrid1.Columns(2).Width = 1600
    DataGrid1.Columns(3).Width = 1600
    DataGrid1.Columns(4).Width = 1600
    DataGrid1.Columns(5).Width = 1550
End Sub
