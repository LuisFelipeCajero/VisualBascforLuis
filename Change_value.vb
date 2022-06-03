Private Sub CommandButton1_Click()

'07 de Abril 2022, Luis Cajero
' Objetivo
' El programa al correr deberá mostrar un formulario, en el cual debería haber los siguientes objetos 8 cajas de texto
' y un botón, el usuario deberá ingresar una palabra o texto en cualquiera de las cajas de texto, y después dar click al botón
' La acción de click del botón deberá pasar el texto a la siguiente caja de texto y borrar el texto donde se encontraba originalmente.

Dim num As Integer
num = 0

If num = 0 Then
            If TextBox1.Value <> "" Then
                TextBox2.Value = TextBox1.Value
                TextBox1.Value = ""
                num = 1
            End If
End If
            
If num = 0 Then
            If TextBox2.Value <> "" Then
                TextBox3.Value = TextBox2.Value
                TextBox2.Value = ""
                num = 1
            End If
End If

If num = 0 Then
            If TextBox3.Value <> "" Then
                TextBox4.Value = TextBox3.Value
                TextBox3.Value = ""
                num = 1
            End If
End If

If num = 0 Then
            If TextBox4.Value <> "" Then
                TextBox5.Value = TextBox4.Value
                TextBox4.Value = ""
                num = 1
            End If
End If

If num = 0 Then
            If TextBox5.Value <> "" Then
                TextBox6.Value = TextBox5.Value
                TextBox5.Value = ""
                num = 1
            End If
End If

If num = 0 Then
            If TextBox6.Value <> "" Then
                TextBox7.Value = TextBox6.Value
                TextBox6.Value = ""
                num = 1
            End If
End If

If num = 0 Then
            If TextBox7.Value <> "" Then
                TextBox8.Value = TextBox7.Value
                TextBox7.Value = ""
                num = 1
            End If
End If
            
If num = 0 Then
            If TextBox8.Value <> "" Then
                TextBox1.Value = TextBox8.Value
                TextBox8.Value = ""
                num = 1
            End If
            
End If



End Sub
