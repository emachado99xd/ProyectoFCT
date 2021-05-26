VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Nuevo"
      Height          =   735
      Left            =   7200
      TabIndex        =   19
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   8760
      TabIndex        =   18
      Top             =   7200
      Width           =   1935
   End
   Begin VB.TextBox txttel 
      Enabled         =   0   'False
      Height          =   495
      Left            =   6600
      TabIndex        =   17
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox txtdir 
      Enabled         =   0   'False
      Height          =   735
      Left            =   6600
      TabIndex        =   16
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtced 
      Enabled         =   0   'False
      Height          =   615
      Left            =   4080
      TabIndex        =   15
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txtape 
      Enabled         =   0   'False
      Height          =   615
      Left            =   4080
      TabIndex        =   14
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtnom 
      Enabled         =   0   'False
      Height          =   615
      Left            =   1200
      TabIndex        =   13
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox txtid 
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   840
      Top             =   7320
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
      Connect         =   "Provider = Microsoft.ACE.OLEDB.12.0; Data Source= "" & App.Path & ""\Base\Base.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider = Microsoft.ACE.OLEDB.12.0; Data Source= "" & App.Path & ""\Base\Base.accdb;Persist Security Info=False"
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
      Left            =   9720
      TabIndex        =   12
      Top             =   1440
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
      Left            =   9960
      TabIndex        =   11
      Top             =   840
      Width           =   90
   End
   Begin VB.Label lblcedula 
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
      Left            =   5760
      TabIndex        =   10
      Top             =   1440
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
      Left            =   6000
      TabIndex        =   9
      Top             =   840
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
      Left            =   1440
      TabIndex        =   8
      Top             =   1440
      Width           =   90
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono:"
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
      Left            =   8280
      TabIndex        =   6
      Top             =   1440
      Width           =   1350
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Dirección:"
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
      Left            =   8280
      TabIndex        =   5
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Cedula:"
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
      Left            =   4560
      TabIndex        =   4
      Top             =   1440
      Width           =   1050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Apellido:"
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
      Left            =   4560
      TabIndex        =   3
      Top             =   840
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
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
      TabIndex        =   2
      Top             =   1440
      Width           =   1050
   End
   Begin VB.Label Label2 
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
      TabIndex        =   1
      Top             =   840
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Factura"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    With RsCliente
        .Requery
        .AddNew
        !Nombre = Val(txtnom.Text)
        !Apellido = Val(txtape.Caption)
        !Cedula = Val(txtced.Text)
        !Direccion = Val(txtdir.Text)
        !Telefono = Val(txttel.Text)
        .UpdateBatch
    End With
    txnom.Enabled = False
    txtape.Enabled = False
    txtced.Enabled = False
    txtdir.Enabled = False
    txttel.Enabled = False
    txnom.Text = ""
    txtape.Text = ""
    txtced.Text = ""
    txtdi.Text = ""
    txttel.Text = ""
    Command1.Enabled = False
    Command2.Enabled = True
End Sub
Private Sub Command2_Click()
    Command2.Enabled = False
    Command1.Enabled = True
    txnom.Enabled = True
    txtape.Enabled = True
    txtced.Enabled = True
    txtdir.Enabled = True
    txttel.Enabled = True
End Sub

Private Sub Form_Load()
 Cliente
End Sub
Private Sub txtid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With RsCliente
            .Requery
            .Find "Id_Cliente='" & Trim(txtid.Text) & "'"
            If .BOF Or .EOF Then Exit Sub
            lblnombre.Caption = !Nombre
            lblapellido.Caption = !Apellido
            lblcedula.Caption = !Cedula
            lbldireccion.Caption = !Direccion
            lbltelefono.Caption = !Telefono
        End With
    End If
End Sub

