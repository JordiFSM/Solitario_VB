Public Class Pila
    Public Property Cartas As List(Of Carta) = New List(Of Carta)
    Public Property Actual As Integer = 0
    Public Property Beginer As Integer = 0

    Public Property LastCartPoint As Point



    'Agrega una carta a la lista de cartas siempre y cuando el maximo sea mayor a la cantidad actual de cartas'
    Public Sub pushDeck(pCart As Carta)
        Cartas.Add(pCart)
        Actual = Actual + 1
    End Sub

    Public Sub popDeck()
        Cartas.RemoveAt(Actual - 1)
        Actual = Actual - 1
    End Sub

    Public Sub defineBegin(pNum As Integer)
        Beginer = pNum
    End Sub

    'Verifica si la ultima carta de la pila y la entrante tienen color diferente'
    Public Function requestDiferentColorAdd(pCart As Carta) As Boolean
        If Actual = 0 Then
            Return True
        ElseIf Cartas(Actual - 1).Color <> pCart.Color Then
            Return True
        End If
        Return False
    End Function
    'Verifica si el numero de la carta es el siguiente de la actual en la pila de tipos'
    Public Function requestMoreNum(pCart As Carta) As Boolean
        If Actual = 0 And pCart.Numero = 1 Then
            Return True
        ElseIf Actual <> 0 Then
            If Cartas(Actual - 1).Numero + 1 = pCart.Numero Then
                Return True
            End If
            Return False
        End If
        Return False
    End Function
    'Verifica si el numero de la carta es el siguiente de la actual en la pila'
    Public Function requestLessNum(pCart As Carta) As Boolean
        If Actual = 0 And pCart.Numero = 13 Then
            Return True
        ElseIf Actual <> 0 Then
            If Cartas(Actual - 1).Numero - 1 = pCart.Numero Then
                Return True
            End If
            Return False
        End If
    End Function

    'Verifica el tipo de la carta actual de la pila con la entrante
    Public Function equalType(pCart As Carta) As Boolean
        If Actual = 0 Then
            Return True
        ElseIf Cartas(Actual - 1).Tipo = pCart.Tipo Then
            Return True
        End If
        Return False
    End Function

    Public Sub removeAllCart()
        Cartas.Clear()
        Actual = 0
    End Sub

    Public Function isPicture(sender As Object) As Boolean
        For Each cart As Carta In Cartas
            If cart.Pic.Location = sender.Location Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Function isPicturePosition(sender As Object) As Integer
        Dim res As Integer = 0
        For Each cart As Carta In Cartas
            If cart.Pic.Location = sender.Location Then
                Return res
            End If
            res += 1
        Next
        Return res
    End Function

    Public Sub agregarCartas(lista As List(Of Carta))
        For Each cart As Carta In lista
            Cartas.Add(cart)
        Next
    End Sub

End Class
