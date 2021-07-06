VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   12600
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3120
      Left            =   600
      TabIndex        =   11
      Top             =   3120
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   5503
      _Version        =   393216
      BackColor       =   -2147483648
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   20
      AllowAddNew     =   -1  'True
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
         Name            =   "Showcard Gothic"
         Size            =   9.75
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
      Left            =   11280
      TabIndex        =   24
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
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
      Left            =   7800
      TabIndex        =   23
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdguardar 
      Caption         =   "Guardar Factura"
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
      Left            =   7680
      TabIndex        =   22
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox txttotal 
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
      Height          =   375
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   7200
      Width           =   1455
   End
   Begin VB.TextBox txtiva 
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
      Height          =   405
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   7560
      Width           =   1455
   End
   Begin VB.TextBox txtsubtotal 
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
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "++"
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
      Left            =   7200
      TabIndex        =   13
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtid 
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
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   13080
      Top             =   4440
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
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Left            =   480
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   10095
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   345
      Left            =   600
      TabIndex        =   18
      Top             =   6720
      Width           =   1560
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Iva:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   345
      Left            =   1680
      TabIndex        =   17
      Top             =   7560
      Width           =   555
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   345
      Left            =   4320
      TabIndex        =   16
      Top             =   7200
      Width           =   1020
   End
   Begin VB.Label lblid 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   14040
      TabIndex        =   15
      Top             =   840
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
      Left            =   14280
      TabIndex        =   14
      Top             =   1440
      Width           =   90
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Productos"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   540
      Left            =   4200
      TabIndex        =   12
      Top             =   2400
      Width           =   2745
   End
   Begin VB.Label lbltelefono 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   6240
      TabIndex        =   10
      Top             =   1560
      Width           =   60
   End
   Begin VB.Label lbldireccion 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9840
      TabIndex        =   9
      Top             =   1200
      Width           =   60
   End
   Begin VB.Label lblapellido 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   6360
      TabIndex        =   8
      Top             =   960
      Width           =   60
   End
   Begin VB.Label lblnombre 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1800
      TabIndex        =   7
      Top             =   1560
      Width           =   60
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   345
      Left            =   4680
      TabIndex        =   5
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   345
      Left            =   8160
      TabIndex        =   4
      Top             =   1200
      Width           =   1590
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cedula:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   345
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   345
      Left            =   4800
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   345
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   1260
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Factura"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   540
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   2025
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim IdProducto, Cantidad As Integer
Private Sub cmdguardar_Click()
Dim Productos As Integer
    With RsFactura
        .Requery
        .Find "Id_Factura='" & Trim(lblfactura.Caption) & "'"
        !fecha = Date
        !Hora = Time
        !Subtotal = Val(txtsubtotal.Text)
        !Iva = Val(txtiva.Text)
        !Total = Val(txttotal.Text)
        !Validar = True
        .UpdateBatch
    End With
    Productos = RsTemporal.RecordCount
    RsTemporal.Requery
    RsTemporal.MoveFirst
    For X = 1 To Productos
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
        If X = Productos Then Else RsTemporal.MoveNext
    Next
    MsgBox "La factura ha sido creada", vbInformation, "Aviso"
    Subtotal = 0
    With RsTemporal
        .Requery
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
    borrar
    limpiar
    DataGrid
    Command1.Enabled = True
    cmdguardar.Enabled = False
    Command2.Enabled = False
    txtid.Locked = False
    txtid.SetFocus
    lblfactura.Caption = ""
    lblid.Caption = ""
End Sub



Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    With RsFactura
     .Requery
     .Find "Id_Factura='" & Trim(lblfactura.Caption) & "'"
     .Delete
     .UpdateBatch
    End With
    borrar
    limpiar
    DataGrid
    Command1.Enabled = True
    Command2.Enabled = False
    txtid.Locked = False
    cmdguardar.Enabled = False
    txtid.SetFocus
End Sub

Private Sub Command3_Click()
    If lblnombre.Caption = "" Then MsgBox "Ingrese un cliente", vbInformation, "Aviso": txtid.SetFocus: Exit Sub
    If RsTemporal.State = 1 Then RsTemporal.Close
    ProductosFactura.Show
End Sub





Private Sub Form_Activate()
    DataGrid
    If txtsubtotal.Text = "" Then
    Else
        cmdguardar.Enabled = True
    End If
End Sub

Private Sub Form_Load()
Form1.Picture = LoadPicture(App.Path & "\IMG\Fondo1.jpg")
Clientes
Factura
Producto
Temporal
DetalleFactura
Set DataGrid1.DataSource = RsTemporal
borrar
cmdguardar.Enabled = False
Command2.Enabled = False
DataGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If RsTemporal.State = 1 Then RsTemporal.Close
    If lblnombre.Caption = "" Then
        Exit Sub
    Else
        With RsFactura
        .Requery
        .Find "Id_Factura='" & Trim(lblfactura.Caption) & "'"
        .Delete
        .UpdateBatch
        End With
        borrar
    End If
End Sub

Private Sub txtid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtid.Text = "" Then Exit Sub
        With RsCliente
            .Requery
            .Find "Cedula='" & Trim(txtid.Text) & "'"
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
        txtid.Locked = True
        Command2.Enabled = True
        Command1.Enabled = False
    End If
End Sub

Sub borrar()
    With RsTemporal
        .Requery
        If .BOF Or .EOF Then Exit Sub
        For X = 1 To .RecordCount
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
    txtid.Text = ""
    lblnombre.Caption = ""
    lblapellido.Caption = ""
    lbltelefono.Caption = ""
    lbldireccion.Caption = ""
    txtsubtotal.Text = ""
    txtiva.Text = ""
    txttotal.Text = ""
End Sub


