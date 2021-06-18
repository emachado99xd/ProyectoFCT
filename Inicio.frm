VERSION 5.00
Begin VB.Form Inicio 
   Caption         =   "Form2"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12675
   LinkTopic       =   "Form2"
   ScaleHeight     =   7170
   ScaleWidth      =   12675
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image5 
      Height          =   1095
      Left            =   6600
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   3495
   End
   Begin VB.Image Image4 
      Height          =   1095
      Left            =   6600
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   3495
   End
   Begin VB.Image Image3 
      Height          =   1095
      Left            =   6600
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   6600
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bienvenido"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   900
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Width           =   4260
   End
   Begin VB.Image Image1 
      Height          =   7215
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Image1.Picture = LoadPicture(App.Path & "\IMG\Fondo_Inicio.jpg")
    Image2.Picture = LoadPicture(App.Path & "\IMG\Botonfactura1.jpg")
    Image3.Picture = LoadPicture(App.Path & "\IMG\Boton1.jpg")
    Image4.Picture = LoadPicture(App.Path & "\IMG\Botoncliente1.jpg")
    Image5.Picture = LoadPicture(App.Path & "\IMG\Botonproducto1.jpg")
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = LoadPicture(App.Path & "\IMG\Botonfactura2.jpg")
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = LoadPicture(App.Path & "\IMG\Botonfactura1.jpg")
    Form1.Show
End Sub
Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image3.Picture = LoadPicture(App.Path & "\IMG\Boton2.jpg")
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image3.Picture = LoadPicture(App.Path & "\IMG\Boton1.jpg")
    Anular.Show
End Sub
Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image4.Picture = LoadPicture(App.Path & "\IMG\Botoncliente2.jpg")
End Sub
Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image4.Picture = LoadPicture(App.Path & "\IMG\Botoncliente1.jpg")
    ModificarCliente.Show
End Sub
Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image5.Picture = LoadPicture(App.Path & "\IMG\Botonproducto2.jpg")
End Sub
Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image5.Picture = LoadPicture(App.Path & "\IMG\Botonproducto1.jpg")
    ModificarProducto.Show
End Sub
