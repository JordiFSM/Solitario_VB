Public Class Tablero
    Public Property Stack1 As Pila = New Pila
    Public Property Stack2 As Pila = New Pila
    Public Property Stack3 As Pila = New Pila
    Public Property Stack4 As Pila = New Pila
    Public Property Stack5 As Pila = New Pila
    Public Property Stack6 As Pila = New Pila
    Public Property Stack7 As Pila = New Pila
    Public Property Deck As Pila = New Pila
    Public Property Content1 As Pila = New Pila
    Public Property Content2 As Pila = New Pila
    Public Property Content3 As Pila = New Pila
    Public Property Content4 As Pila = New Pila
    ' Define la cantidad de cartas iniciales de los decks
    Public Sub settingBoardBegin()
        Stack1.defineBegin(1)
        Stack2.defineBegin(2)
        Stack3.defineBegin(3)
        Stack4.defineBegin(4)
        Stack5.defineBegin(5)
        Stack6.defineBegin(6)
        Stack7.defineBegin(7)
        Deck.defineBegin(24)

        insertAllCartsonDeck()
        shuffleDeck()
        'configAllstacks()

    End Sub
    'Crea todas las 52 cartas y las inserta en el deck 
    Public Sub insertAllCartsonDeck()
        Dim Num As Integer = 1
        Dim Tipes As New List(Of String)
        Tipes.Add("Corazones")
        Tipes.Add("Diamantes")
        Tipes.Add("Picas")
        Tipes.Add("Treboles")
        For j = 0 To 3
            For i = 1 To 13
                Dim Cart As New Carta
                Cart.defineCart(Tipes(j), i)
                Deck.pushDeck(Cart)
            Next
        Next
    End Sub

    'Baraja de forma aleatoria el deck
    Public Sub shuffleDeck()
        For i = 0 To Deck.Actual
            Dim index As Integer
            Dim CartTemporal As New Carta
            index = CInt(New Random(CInt(Date.Now.Ticks And Integer.MaxValue)).Next Mod 52)
            CartTemporal = Deck.Cartas(i)

            Deck.Cartas(i) = Deck.Cartas(index)
            Deck.Cartas(index) = CartTemporal
        Next
    End Sub
    'Genera las cartas iniciales de los 7 stacks de juego dejando el deck solo con 24
    Public Sub configAllstacks()
        For i = 0 To Stack1.Beginer - 1
            Dim cartTemp As New Carta
            cartTemp = Deck.Cartas(Deck.Actual)

            Stack1.pushDeck(cartTemp)
            Deck.popDeck()
        Next
        For i = 0 To Stack2.Beginer - 1
            Dim cartTemp As New Carta
            cartTemp = Deck.Cartas(Deck.Actual)

            Stack2.pushDeck(cartTemp)
            Deck.popDeck()
        Next
        For i = 0 To Stack3.Beginer - 1
            Dim cartTemp As New Carta
            cartTemp = Deck.Cartas(Deck.Actual)

            Stack2.pushDeck(cartTemp)
            Deck.popDeck()
        Next
        For i = 0 To Stack4.Beginer - 1
            Dim cartTemp As New Carta
            cartTemp = Deck.Cartas(Deck.Actual)

            Stack2.pushDeck(cartTemp)
            Deck.popDeck()
        Next
        For i = 0 To Stack5.Beginer - 1
            Dim cartTemp As New Carta
            cartTemp = Deck.Cartas(Deck.Actual)

            Stack2.pushDeck(cartTemp)
            Deck.popDeck()
        Next
        For i = 0 To Stack6.Beginer - 1
            Dim cartTemp As New Carta
            cartTemp = Deck.Cartas(Deck.Actual)

            Stack2.pushDeck(cartTemp)
            Deck.popDeck()
        Next
        For i = 0 To Stack7.Beginer - 1
            Dim cartTemp As New Carta
            cartTemp = Deck.Cartas(Deck.Actual)

            Stack2.pushDeck(cartTemp)
            Deck.popDeck()
        Next
    End Sub






End Class
