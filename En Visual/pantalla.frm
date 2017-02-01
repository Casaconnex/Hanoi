VERSION 5.00
Begin VB.Form pantalla 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Torres de Hanoi"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   11595
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4800
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   480
      Picture         =   "pantalla.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   10185
   End
   Begin VB.Shape PALO 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   3255
      Index           =   2
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape PALO 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   3255
      Index           =   3
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape PALO 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   3255
      Index           =   1
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "No. de Movimientos"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   4
      Top             =   7200
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9240
      TabIndex        =   3
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label torre3 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label torre2 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label torre1 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image base1 
      Height          =   300
      Left            =   960
      Picture         =   "pantalla.frx":1D5E8
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   2805
   End
   Begin VB.Shape aros 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   600
      Shape           =   2  'Oval
      Top             =   1560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image BASE2 
      Height          =   300
      Left            =   4200
      Picture         =   "pantalla.frx":1F206
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2805
   End
   Begin VB.Image BASE3 
      Height          =   300
      Left            =   3360
      Picture         =   "pantalla.frx":20E24
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   2805
   End
End
Attribute VB_Name = "pantalla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If MOV Then
    CONT = CONT + 1
    Label1 = CONT
    HANOI_IMPAR
End If

End Sub

Private Sub Form_Load()

CARGAR_COLORES

n = InputBox("ingrese la cantidad de elementos", "TORRES DE HANOI", 1)
If n <= 0 Or n > 17 Then
    MsgBox "EL VALOR " & n & " ESTA FUERA DE RANGO ", vbExclamation, "TORRES DE HANOI"
    n = InputBox("ingrese la cantidad de elementos", "TORRES DE HANOI", 1)
End If

'DIMENSIONAR
CREAR_AROS
'CREAR_CAJAS
inicializar
MOSTRAR

End Sub


Public Sub CREAR_AROS()
    PALO(1).Width = 50
    PALO(1).Visible = True
    PALO(1).Height = 250 * n + 400
    base1.Left = PALO(1).Left - base1.Width / 2 + PALO(1).Width / 2
    base1.Top = PALO(1).Top + PALO(1).Height - 200
    
    PALO(2).Height = PALO(1).Height
    PALO(2).Width = PALO(1).Width
    PALO(2).Top = PALO(1).Top
    PALO(2).Left = PALO(1).Left + 3700
    PALO(2).Visible = True
    BASE2.Left = PALO(2).Left - BASE2.Width / 2 + PALO(2).Width / 2
    BASE2.Top = PALO(2).Top + PALO(2).Height - 200
    
    PALO(3).Height = PALO(1).Height
    PALO(3).Width = PALO(1).Width
    PALO(3).Top = PALO(1).Top
    PALO(3).Left = PALO(2).Left + 3700
    PALO(3).Visible = True
    BASE3.Left = PALO(3).Left - BASE3.Width / 2 + PALO(3).Width / 2
    BASE3.Top = PALO(3).Top + PALO(3).Height - 200
    
    For X = 1 To n
        Load aros(X)
        With aros(X)
            .FillColor = colores(X)
            .Width = X * 150
            .Height = 250
            .Left = PALO(1).Left - .Width / 2 + PALO(1).Width / 2
            .Top = 100 + PALO(1).Top + (X - 1) * .Height
            .Visible = True
        End With
    Next X
    PALO(1).Visible = True
End Sub
Public Sub CREAR_CAJAS()
    For X = 1 To n
        
        Load torre1(X)
        Load torre2(X)
        Load torre3(X)
        
        With torre1(X)
            .Top = .Height + torre1(X - 1).Top
            .Visible = True
            
        End With
        
        With torre2(X)
            .Top = .Height + torre2(X - 1).Top
            .Visible = True
        End With
        
        With torre3(X)
            .Top = .Height + torre3(X - 1).Top
            .Visible = True
        End With
        
    Next X
End Sub

Public Sub DIMENSIONAR()
    pantalla.Height = n * torre1(0).Height + 1000

    
End Sub

Private Sub Timer1_Timer()
If MOV Then
    CONT = CONT + 1
    Label1 = CONT
    If ESPAR(n) Then
        HANOI_PAR
    Else
        HANOI_IMPAR
    End If
End If

End Sub
