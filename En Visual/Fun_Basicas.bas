Attribute VB_Name = "FuncionesBásicas"
Public Sub inicializar()
    I = 1
    K = n + 1
    J = n + 1
    
    ReDim t1(1 To n)
    ReDim t2(1 To n)
    ReDim t3(1 To n)
    
    For X = 1 To n
        t1(X) = X
        t2(X) = 0
        t3(X) = 0
    Next X
    DESDE = 1
    VALIDO = True
    MOV = True
    CONT = CONT + 1
End Sub
Public Sub MOSTRAR()
    For X = 1 To n
        pantalla.aros(X).Visible = False
    Next X
    
    Dim ELEMENTOS As Integer
    ELEMENTOS = 0
    
    For X = n To 1 Step -1
        
        If t1(X) <> 0 Then
            pantalla.aros(t1(X)).Top = pantalla.PALO(1).Top + pantalla.PALO(1).Height - 600 - ELEMENTOS * pantalla.aros(t1(X)).Height
            pantalla.aros(t1(X)).Left = pantalla.PALO(1).Left - pantalla.aros(t1(X)).Width / 2 + pantalla.PALO(1).Width / 2
            pantalla.aros(t1(X)).Visible = True
            ELEMENTOS = ELEMENTOS + 1
        End If
    Next X
    
    ELEMENTOS = 0
    For X = n To 1 Step -1
        
        If t2(X) <> 0 Then
            pantalla.aros(t2(X)).Top = pantalla.PALO(2).Top + pantalla.PALO(2).Height - 600 - ELEMENTOS * pantalla.aros(t2(X)).Height
            pantalla.aros(t2(X)).Left = pantalla.PALO(2).Left - pantalla.aros(t2(X)).Width / 2 + pantalla.PALO(2).Width / 2
            pantalla.aros(t2(X)).Visible = True
            ELEMENTOS = ELEMENTOS + 1
        End If
    Next X
    
    ELEMENTOS = 0
    For X = n To 1 Step -1
        
        If t3(X) <> 0 Then
            pantalla.aros(t3(X)).Top = pantalla.PALO(3).Top + pantalla.PALO(3).Height - 600 - ELEMENTOS * pantalla.aros(t3(X)).Height
            pantalla.aros(t3(X)).Left = pantalla.PALO(3).Left - pantalla.aros(t3(X)).Width / 2 + pantalla.PALO(3).Width / 2
            pantalla.aros(t3(X)).Visible = True
            ELEMENTOS = ELEMENTOS + 1
        End If
    Next X
    
    Exit Sub
    
End Sub


Public Function ESPAR(NUMERO As Integer) As Boolean
    If NUMERO Mod 2 = 0 Then
        ESPAR = True
    Else
        ESPAR = False
    End If
End Function


Public Sub CARGAR_COLORES()
    ReDim colores(1 To 17)
    
    colores(1) = RGB(244, 70, 20)
    colores(2) = RGB(233, 142, 75)
    colores(3) = RGB(235, 190, 73)
    colores(4) = RGB(252, 248, 67)
    colores(5) = RGB(173, 252, 67)
    colores(6) = RGB(82, 245, 139)
    colores(7) = RGB(82, 245, 221)
    colores(8) = RGB(82, 217, 245)
    colores(9) = RGB(83, 168, 244)
    colores(10) = RGB(98, 143, 230)
    colores(11) = RGB(104, 98, 230)
    colores(12) = RGB(203, 194, 150)
    colores(13) = RGB(134, 171, 193)
    colores(14) = RGB(170, 170, 170)
    colores(15) = RGB(255, 255, 255)
    colores(16) = RGB(0, 0, 0)
End Sub
