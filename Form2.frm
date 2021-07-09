VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10710
   LinkTopic       =   "Form2"
   ScaleHeight     =   5535
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   600
      List            =   "Form2.frx":000D
      TabIndex        =   5
      Text            =   "Opciones"
      Top             =   2880
      Width           =   2655
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
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   1920
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
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
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


Private Sub Combo1_Click()
    If Combo1.Text = "Fecha" Then
        Text2.Visible = True
        Text3.Visible = True
    Else
        If Combo1.Text = "Precio" Then
        Text2.Visible = True
        End If
    End If
End Sub

Private Sub Command1_Click()
    If Combo1.Text = "Fecha" Then
        With RsFactura
            If .State = 1 Then .Close
            Dim fecha, Inicio As Date
            Inicio = Text2.Text
            fecha = Text3.Text
            a = "#" & Inicio & "#"
            .Open "Select * From Factura Where ((Factura.[Fecha])= " & a & ")", Base, adOpenStatic, adLockBatchOptimistic
            If .EOF Or .BOF Then MsgBox "Fecha(a) no hay": Exit Sub
                b = "#" & fecha & "#"
            If .State = 1 Then .Close
                .Open "Select * From Factura Where ((Factura.[Fecha])= " & b & ")", Base, adOpenStatic, adLockBatchOptimistic
            If .EOF Or .BOF Then MsgBox "Fecha(b) no hay": Exit Sub
            If .State = 1 Then .Close
                .Open "Select * From Factura Where ((Factura.[Fecha])>= " & a & ") AND ((Factura.[Fecha])<= " & b & ")", Base, adOpenStatic, adLockBatchOptimistic
            Set DataReport2.DataSource = RsFactura
            DataReport2.Show
        End With
    Else
        If Combo1.Text = "Anuladas" Then
            With RsFactura
                If .State = 1 Then .Close
                '.Open "Select * From Factura Where ((Factura.[Fecha])>= " & a & ") AND ((Factura.[Fecha])<= " & b & ")", Base, adOpenStatic, adLockBatchOptimistic
                .Open "Select * From Factura Where ((Factura.[Validar])= False)", Base, adOpenStatic, adLockBatchOptimistic
                Set DataReport2.DataSource = RsFactura
                DataReport2.Show
            End With
        Else
            If Combo1.Text = "Precio" Then
                With RsFactura
                    a = Val(Text2.Text)
                    If .State = 1 Then .Close
                    .Open "Select * From Factura Where ((Factura.[Total])> " & a & ")", Base, adOpenStatic, adLockBatchOptimistic
                    Set DataReport2.DataSource = RsFactura
                    DataReport2.Show
                End With
            End If
        End If
    End If
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

