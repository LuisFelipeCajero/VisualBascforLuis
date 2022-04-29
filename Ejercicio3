'Cambio de un valor en una celda de 10000 valores 

For x = 20 To 10000
 mi_caracter = Mid(Cells(15, 10).Value, x, 1)
     
      If Cells(x, 1).Value = "Thunderstruck" Then
            
             mi_numero = Len(Cells(x, 1).Value)
            
             For y = 1 To mi_numero
                 mi_numero = Mid(Cells(x, 1).Value, y, 1)

                 Dim volteado As String
                 volteado = mi_numero & volteado
             Next y
            
             MsgBox volteado
        
            
     End If
     
     Next x
