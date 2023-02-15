Private Sub CommandButton2_Click()
'Hacer un programa que cambie los precios
'Boton nuevo archivo de salida
'Localizar la columna de precios nuevos
'Cambiar los valores de los precios anteriores por los valores de precios nuevos
'Declarar una variable donde almacene el valor de la celda "precio nuevo"

ruta_nueva = "D:\Paso\nuevos_precios.txt"
ruta_salida = "D:\Paso\nuevos_salida.txt"
 
 Dim cuenta As Integer
 
 
Open ruta_nueva For Input As #1
Open ruta_salida For Output As #2

  
   
Do Until EOF(1)
    
    Line Input #1, textline
    var1 = textline
    var2 = Len(var1)
    
    For y = 1 To var2
    var3 = Mid(var1, y, 1)
        If var3 = "a" Then
        cuenta = cuenta + 1
        End If
        
 
    
    Next y
    
    Write #2, cuenta
    cuenta = 0
    
    
    
    
Loop
Close #2
Close #1

   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
'For x = 2 To 14

        'precio_nuevo = Cells(x, 17).Value
        'If precio_nuevo < Cells(x, 16).Value Or precio_nuevo > Cells(x, 16).Value Then
        'Cells(x, 16).Value = precio_nuevo
        'Else
        'MsgBox ("No hay valores que cambiar en este producto")
        'End If
'Next x
 
 
 
 
 
 'For x = 1 To 10
    'Dim tabla As Integer
    'tabla = x * 1
    ' MsgBox "1 " & "X " & x & " =" & " " & tabla
    
    'Cells(x, 2).Value = "1 " & "X " & x & " =" & " " & tabla
'Next x
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
    'Dim precio_nuevo As Integer
    'x = 1
    
    'Cells(15, x).Value = precio_nuevo
    'x = x + 1
    'if precio_nuevo ><




End Sub
