Option Explicit
Dim AClaves() As String
Dim impr As String

Sub buscarClaves()
    impr = ""
    Erase AClaves()
    ReDim Preserve AClaves(10)

    Dim cont As Integer
    cont = 0
    Dim nClave As String
    Dim iRet As Integer  

    Do While nClave <> "n" ' Presionar n para terminar de capturar
        nClave = InputBox("Ingrese Clave")
        If bClaves(nClave, AClaves()) Then
            ' Aca irá condicion para registrar duplicado o no.
                 iRet = MsgBox("Clave duplicado - ¿Desea registrarlo?", vbYesNo, "Registrar claves.")
                 If iRet = vbYes Then
                    If cont > 10 Then
                        ' Si rebasa el indice inicial(en este caso 10 elementos) redimensionar arreglo de 1 en 1
                            ReDim Preserve AClaves(cont)
                    End If
                    AClaves(cont) = nClave
                    cont = cont + 1
                 End If
            Else

                If cont > 10 Then
                MsgBox ("ReDim")
                    ReDim Preserve AClaves(cont)
                End If
                AClaves(cont) = nClave
                cont = cont + 1
              End If
        Loop

    ' Imprimir datos ingresados... Opcional
    Dim nnclave As Variant
    Dim n As Integer
    For n = 0 To UBound(AClaves)
        If AClaves(n) <> "" Then
            impr = impr & " " & AClaves(n)
        End If
    Next n
    MsgBox (impr)
    End Sub

    Function bClaves(ByVal clve As String, ByRef Datos() As String) As Boolean
        bClaves = False
        Dim clave As Variant ' Peticion del VBA
        For Each clave In Datos
            If clave <> "" Then
                If clave = clve Then
                    bClaves = True
                    Exit Function
                End If
            Else
                bClaves = False
                Exit For
            End If
        Next
    End Function
