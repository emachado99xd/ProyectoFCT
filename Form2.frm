VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10815
   LinkTopic       =   "Form2"
   ScaleHeight     =   5160
   ScaleWidth      =   10815
   StartUpPosition =   3  'Windows Default
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
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   1920
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
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   7800
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Producto"
      Height          =   735
      Left            =   6840
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Factura"
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    With RsFactura
        If .State = 1 Then .Close
        Dim fecha, inicio As Date
        inicio = Text2.Text
        fecha = Text3.Text
        a = "#" & inicio & "#"
        .Open "Select * From Factura Where ((Factura.[Fecha])= " & a & ")", Base, adOpenStatic, adLockBatchOptimistic
        If .EOF Or .BOF Then MsgBox "Fecha(a) no hay": Exit Sub
            b = "#" & fecha & "#"
        If .State = 1 Then .Close
            .Open "Select * From Factura Where ((Factura.[Fecha])= " & b & ")", Base, adOpenStatic, adLockBatchOptimistic
        If .EOF Or .BOF Then MsgBox "Fecha(b) no hay": Exit Sub
        If .State = 1 Then .Close
            .Open "Select * From Factura Where ((Factura.[Fecha])>= " & a & ") AND ((Factura.[Fecha])<= " & b & ")", Base, adOpenStatic, adLockBatchOptimistic
        
        '.Open "SELECT *FROM Factura Where [Fecha]BETWEEN '" & inicio & "' and '" & fecha & "'"
        Set DataReport2.DataSource = RsFactura
        DataReport2.Show
    End With
End Sub

Private Sub Command2_Click()
    With RsProductos
        If .State = 1 Then .Close
        Dim buscar As String
        buscar = Val(Text1.Text)
        .Open "SELECT *FROM Producto Where [Stock] <= '" & buscar & "'"
        Set DataReport3.DataSource = RsProductos
        DataReport3.Show
    End With
End Sub

Private Sub Form_Load()
    Factura
    Producto
End Sub

