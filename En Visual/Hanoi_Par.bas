Attribute VB_Name = "HanoiPAR"

Public Sub HANOI_PAR()
If I = n + 1 And J = n + 1 Then
    
    MsgBox "TERMINO EL JUEGO", vbCritical, "TORRES DE HANOI"
    MOV = False

Else

Select Case (DESDE)
    Case 1:     Select Case VALIDO
                        Case True:  MOVTORRE1PAR
                        Case False:  Select Case HASTA
                                                Case 2:      VALIDO = True
                                                                 MOVTORRE3PAR
                                                Case 3:     VALIDO = True
                                                                MOVTORRE2PAR
                                            End Select
                    End Select
    
                    
                    
                    
                    
    Case 2:     Select Case VALIDO
                        Case True: MOVTORRE2PAR
                        Case False:  Select Case HASTA
                                            Case 1:   VALIDO = True
                                                            MOVTORRE3PAR
                                            Case 3:     VALIDO = True
                                                            MOVTORRE1PAR
                                            End Select
                        End Select
                    
    Case 3:     Select Case VALIDO
                        Case True:  MOVTORRE3PAR
                        Case False:  Select Case HASTA
                                                Case 2:      VALIDO = True
                                                                 MOVTORRE1PAR
                                                Case 1:     VALIDO = True
                                                                MOVTORRE2PAR
                                            End Select
                    End Select
    
End Select

'HANOI_PAR
End If




End Sub

Public Sub MOVTORRE1PAR()
    
    If (I = n + 1) Then
        
        VALIDO = False
        
    Else
        
        If ESPAR(I) = False Then
            
            If (J = n + 1) Then
                
                J = J - 1
                t2(J) = t1(I)
                t1(I) = 0
                I = I + 1
                DESDE = 1
                HASTA = 2
                VALIDO = True
                MOSTRAR
            
            Else
                If t1(I) < t2(J) Then
                    J = J - 1
                    t2(J) = t1(I)
                    t1(I) = 0
                    I = I + 1
                    DESDE = 1
                    HASTA = 2
                    VALIDO = True
                    MOSTRAR
                
                Else
                    VALIDO = False
                    DESDE = 1
                End If
                
            End If
         Else
         If (K = n + 1) Then
                
                K = K - 1
                t3(K) = t1(I)
                t1(I) = 0
                I = I + 1
                DESDE = 1
                HASTA = 3
                VALIDO = True
                MOSTRAR
            
            Else
                If t1(I) < t3(K) Then
                    K = K - 1
                    t3(K) = t1(I)
                    t1(I) = 0
                    I = I + 1
                    DESDE = 1
                    HASTA = 3
                    VALIDO = True
                    MOSTRAR
                    Exit Sub
                
                Else
                    VALIDO = False
                    DESDE = 1
                End If
            End If
        End If
        
    End If
End Sub

Public Sub MOVTORRE2PAR()
    
    If (J = n + 1) Then
        
        VALIDO = False
        
    Else
        
        If ESPAR(J) = False Then
            
            If (I = n + 1) Then
                
                I = I - 1
                t1(I) = t2(J)
                t2(J) = 0
                J = J + 1
                HASTA = 1
                DESDE = 2
                VALIDO = True
                MOSTRAR
            
            Else
                If t1(I) > t2(J) Then
                    I = I - 1
                    t1(I) = t2(J)
                    t2(J) = 0
                    J = J + 1
                    HASTA = 1
                    DESDE = 2
                    VALIDO = True
                    MOSTRAR
                    
                Else
                    VALIDO = False
                    DESDE = 2
                End If
                
            End If
         Else
         If (K = n + 1) Then
                
                K = K - 1
                t3(K) = t2(J)
                t2(J) = 0
                J = J + 1
                HASTA = 3
                DESDE = 2
                VALIDO = True
                MOSTRAR
                
            Else
                If t2(J) < t3(K) Then
                    K = K - 1
                    t3(K) = t2(J)
                    t2(J) = 0
                    J = J + 1
                    HASTA = 3
                    DESDE = 2
                    VALIDO = True
                    MOSTRAR
                
                Else
                    VALIDO = False
                    DESDE = 2
                End If
            End If
        End If
        
    End If
End Sub

Public Sub MOVTORRE3PAR()
    
    If (K = n + 1) Then
        
        VALIDO = False
        
    Else
        
        If ESPAR(K) = False Then
            
            If (I = n + 1) Then
                
                I = I - 1
                t1(I) = t3(K)
                t3(K) = 0
                K = K + 1
                HASTA = 1
                VALIDO = True
                DESDE = 3
                MOSTRAR
            
            Else
                If t3(K) < t1(I) Then
                    I = I - 1
                    t1(I) = t3(K)
                    t3(K) = 0
                    K = K + 1
                    HASTA = 1
                    DESDE = 3
                    VALIDO = True
                    
                    MOSTRAR
                
                Else
                    VALIDO = False
                    DESDE = 3
                End If
                
            End If
         Else
         
         If (J = n + 1) Then
                
                K = K - 1
                t2(J) = t3(K)
                t3(K) = 0
                K = K + 1
                HASTA = 2
                DESDE = 3
                VALIDO = True
                MOSTRAR
            
            Else
                If t3(K) < t2(J) Then
                    J = J - 1
                    t2(J) = t3(K)
                    t3(K) = 0
                    K = K + 1
                    HASTA = 2
                    DESDE = 3
                    VALIDO = True
                    MOSTRAR
                
                Else
                    VALIDO = False
                    DESDE = 3
                End If
            End If
        End If
        
    End If
End Sub
