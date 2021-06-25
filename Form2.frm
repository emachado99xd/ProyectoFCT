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
        Dim fecha As Date
        fecha = Date
        .Open "SELECT *FROM Factura Where [Fecha] <= fecha "
        Set DataReport2.DataSource = RsFactura
        DataReport2.Show
    End With
End Sub

Private Sub Form_Load()
    Factura
    
End Sub

