Public Class Form1


    Private isDragging As Boolean
    Public startPoint As Point
    Public sum1 As Integer = 0
    Public cartasMove As Integer = 0

    Public listPictures As New List(Of Image)
    Public listCartsRemoAdd As New List(Of Carta)

    Public cartaComparativa As Carta
    Public cartaComparativaLocation As Point


    Public Property Deck As Pila = New Pila
    Public Property deck2 As Pila = New Pila
    Public Property stack1 As Pila = New Pila
    Public Property stack2 As Pila = New Pila
    Public Property stack3 As Pila = New Pila
    Public Property stack4 As Pila = New Pila
    Public Property stack5 As Pila = New Pila
    Public Property stack6 As Pila = New Pila
    Public Property stack7 As Pila = New Pila


    Public Property content1 As Pila = New Pila
    Public Property content2 As Pila = New Pila
    Public Property content3 As Pila = New Pila
    Public Property content4 As Pila = New Pila



    Public Property stackActual As Pila = New Pila


    Public Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        settingBoardBegin()
        configAllstacks()
        AxWindowsMediaPlayer2.URL = Application.StartupPath() + "\Cartas\AmbienteJazz.mp3"
    End Sub

    Private Sub createListOfPictures()
        listPictures.Clear()
        For i = 0 To cartasMove - 1
            listPictures.Add(stackActual.Cartas(stackActual.Actual - 1 - i).Pic.Image)
        Next
    End Sub

    Private Sub createListOfCartsToRemoveOrAdd()
        listCartsRemoAdd.Clear()
        For i = 0 To cartasMove - 1
            listCartsRemoAdd.Add(stackActual.Cartas(stackActual.Actual - 1 - i))
        Next
    End Sub

    Private Sub PictureBox1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        cartaComparativa = New Carta
        If sender.Location = stack1.LastCartPoint Then
            sender.BringToFront()
            cartasMove = 1
            stackActual = stack1
        ElseIf sender.Location = stack2.LastCartPoint Then
            sender.BringToFront()
            cartasMove = 1
            stackActual = stack2
        ElseIf sender.Location = stack3.LastCartPoint Then
            sender.BringToFront()
            cartasMove = 1
            stackActual = stack3
        ElseIf sender.Location = stack4.LastCartPoint Then
            sender.BringToFront()
            cartasMove = 1
            stackActual = stack4
        ElseIf sender.Location = stack5.LastCartPoint Then
            sender.BringToFront()
            cartasMove = 1
            stackActual = stack5
        ElseIf sender.Location = stack6.LastCartPoint Then
            sender.BringToFront()
            cartasMove = 1
            stackActual = stack6
        ElseIf sender.Location = stack7.LastCartPoint Then
            sender.BringToFront()
            cartasMove = 1
            stackActual = stack7
        ElseIf sender.Location = content1.LastCartPoint Then
            sender.BringToFront()
            cartasMove = 1
            stackActual = content1
        ElseIf sender.Location = content2.LastCartPoint Then
            sender.BringToFront()
            cartasMove = 1
            stackActual = content2
        ElseIf sender.Location = content3.LastCartPoint Then
            sender.BringToFront()
            cartasMove = 1
            stackActual = content3
        ElseIf sender.Location = content4.LastCartPoint Then
            sender.BringToFront()
            cartasMove = 1
            stackActual = content4
        ElseIf sender.Location = deck2.LastCartPoint Then
            sender.BringToFront()
            cartasMove = 1
            stackActual = deck2
        Else
            If stack1.isPicture(sender) Then
                cartasMove = stack1.Actual - stack1.isPicturePosition(sender)
                stackActual = stack1
                TextBox1.Text = "1  "
            ElseIf stack2.isPicture(sender) Then
                cartasMove = stack2.Actual - stack2.isPicturePosition(sender)
                stackActual = stack2
                TextBox1.Text = "2  "
            ElseIf stack3.isPicture(sender) Then
                cartasMove = stack3.Actual - stack3.isPicturePosition(sender)
                stackActual = stack3
                TextBox1.Text = "3  "
            ElseIf stack4.isPicture(sender) Then
                cartasMove = stack4.Actual - stack4.isPicturePosition(sender)
                stackActual = stack4
                TextBox1.Text = "4  "
            ElseIf stack5.isPicture(sender) Then
                cartasMove = stack5.Actual - stack5.isPicturePosition(sender)
                stackActual = stack5
                TextBox1.Text = "5  "
            ElseIf stack6.isPicture(sender) Then
                cartasMove = stack6.Actual - stack6.isPicturePosition(sender)
                stackActual = stack6
                TextBox1.Text = "6  "
            ElseIf stack7.isPicture(sender) Then
                cartasMove = stack7.Actual - stack7.isPicturePosition(sender)
                stackActual = stack7
                TextBox1.Text = "7  "
            End If
            createListOfPictures()
            createListOfCartsToRemoveOrAdd()
            cartaComparativa = listCartsRemoAdd(cartasMove - 1)
            cartaComparativaLocation = cartaComparativa.Pic.Location
            Dim a As Integer = 0
            For i = 0 To cartasMove - 2
                Me.Controls.Remove(listCartsRemoAdd(i).Pic)
                stackActual.Cartas.Remove(listCartsRemoAdd(i))
                stackActual.Actual -= 1
                a += 1
            Next
            stackActual.Cartas.Remove(listCartsRemoAdd(a))
            stackActual.Actual -= 1
            stackActual.LastCartPoint = New Point(cartaComparativa.Pic.Location.X, cartaComparativa.Pic.Location.Y)
        End If
        startPoint = DirectCast(sender, PictureBox).PointToScreen(New Point(e.X, e.Y))
        isDragging = True
    End Sub
    Private Sub PictureBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        If isDragging Then
            If cartasMove = 1 Then
                Dim pb1 As PictureBox = DirectCast(sender, PictureBox)
                Dim EndPoint As Point = pb1.PointToScreen(New Point(e.X, e.Y))
                pb1.Left += (EndPoint.X - startPoint.X)
                pb1.Top += (EndPoint.Y - startPoint.Y)
                startPoint = EndPoint
            Else
                Dim pb2 As PictureBox = pictureBox_Paint()
                Dim pb1 As New PictureBox
                pb1 = DirectCast(sender, PictureBox)
                pb1.BringToFront()
                pb1.Size = pb2.Size
                pb1.Image = pb2.Image
                Dim EndPoint As Point = pb1.PointToScreen(New Point(e.X, e.Y))
                pb1.Left += (EndPoint.X - startPoint.X)
                pb1.Top += (EndPoint.Y - startPoint.Y)
                startPoint = EndPoint
            End If
        End If
    End Sub


    ' Crea un picture box con la lista de imagenes que hay debajo del picture seleccionado
    Public Function pictureBox_Paint() As PictureBox
        Dim picRes As New PictureBox
        picRes.Size = New Size(90, 130 + 20 * (cartasMove - 1))
        Dim bMap As New Drawing.Bitmap(90, 130 + 20 * cartasMove)
        Dim Grf As Graphics = Graphics.FromImage(bMap)
        For i = 0 To cartasMove - 1
            Dim rect As New Rectangle(0, 20 * i, 90, 150)
            Grf.DrawImage(listPictures(cartasMove - 1 - i), rect, 0, 0, 100, 150, CType(2, GraphicsUnit))
            picRes.Image = bMap
        Next
        Return picRes
    End Function

    Private Sub PictureBox1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        isDragging = False
        If cartasMove = 1 Then
            Dim p = DirectCast(sender, PictureBox)
            If stackActual.LastCartPoint <> content1.LastCartPoint And p.Location.Y >= content1.LastCartPoint.Y - 30 And p.Location.Y <= content1.LastCartPoint.Y + 150 And
            p.Location.X >= content1.LastCartPoint.X - 20 And p.Location.X <= content1.LastCartPoint.X + 70 Then
                If content1.equalType(stackActual.Cartas(stackActual.Actual - 1)) And content1.requestMoreNum(stackActual.Cartas(stackActual.Actual - 1)) Then
                    p.Location = New Point(content1.LastCartPoint().X, content1.LastCartPoint().Y)
                    content1.pushDeck(stackActual.Cartas(stackActual.Actual - 1))
                    stackActual.popDeck()
                    If stackActual.LastCartPoint <> deck2.LastCartPoint And stackActual.LastCartPoint <> content2.LastCartPoint And stackActual.LastCartPoint <> content3.LastCartPoint And stackActual.LastCartPoint <> content4.LastCartPoint Then
                        stackActual.LastCartPoint = New Point(stackActual.LastCartPoint.X, stackActual.LastCartPoint.Y - 20)
                    End If
                    verificarLastPosDeck()
                    validarGane()
                Else
                    sender.Location = stackActual.LastCartPoint()
                End If
            ElseIf stackActual.LastCartPoint <> content2.LastCartPoint And p.Location.Y >= content2.LastCartPoint.Y - 30 And p.Location.Y <= content2.LastCartPoint.Y + 150 And
            p.Location.X >= content2.LastCartPoint.X - 20 And p.Location.X <= content2.LastCartPoint.X + 70 Then
                If content2.equalType(stackActual.Cartas(stackActual.Actual - 1)) And content2.requestMoreNum(stackActual.Cartas(stackActual.Actual - 1)) Then
                    p.Location = New Point(content2.LastCartPoint().X, content2.LastCartPoint().Y)
                    content2.pushDeck(stackActual.Cartas(stackActual.Actual - 1))
                    stackActual.popDeck()
                    If stackActual.LastCartPoint <> deck2.LastCartPoint And stackActual.LastCartPoint <> content1.LastCartPoint And stackActual.LastCartPoint <> content3.LastCartPoint And stackActual.LastCartPoint <> content4.LastCartPoint Then
                        stackActual.LastCartPoint = New Point(stackActual.LastCartPoint.X, stackActual.LastCartPoint.Y - 20)
                    End If
                    verificarLastPosDeck()
                    validarGane()
                Else
                    sender.Location = stackActual.LastCartPoint()
                End If
            ElseIf stackActual.LastCartPoint <> content3.LastCartPoint And p.Location.Y >= content3.LastCartPoint.Y - 30 And p.Location.Y <= content3.LastCartPoint.Y + 150 And
            p.Location.X >= content3.LastCartPoint.X - 20 And p.Location.X <= content3.LastCartPoint.X + 70 Then
                If content3.equalType(stackActual.Cartas(stackActual.Actual - 1)) And content3.requestMoreNum(stackActual.Cartas(stackActual.Actual - 1)) Then
                    p.Location = New Point(content3.LastCartPoint().X, content3.LastCartPoint().Y)
                    content3.pushDeck(stackActual.Cartas(stackActual.Actual - 1))
                    stackActual.popDeck()
                    If stackActual.LastCartPoint <> deck2.LastCartPoint And stackActual.LastCartPoint <> content1.LastCartPoint And stackActual.LastCartPoint <> content2.LastCartPoint And stackActual.LastCartPoint <> content4.LastCartPoint Then
                        stackActual.LastCartPoint = New Point(stackActual.LastCartPoint.X, stackActual.LastCartPoint.Y - 20)
                    End If
                    validarGane()
                    verificarLastPosDeck()
                Else
                    sender.Location = stackActual.LastCartPoint()
                End If
            ElseIf stackActual.LastCartPoint <> content4.LastCartPoint And p.Location.Y >= content4.LastCartPoint.Y - 30 And p.Location.Y <= content4.LastCartPoint.Y + 150 And
            p.Location.X >= content4.LastCartPoint.X - 20 And p.Location.X <= content4.LastCartPoint.X + 70 Then
                If content4.equalType(stackActual.Cartas(stackActual.Actual - 1)) And content4.requestMoreNum(stackActual.Cartas(stackActual.Actual - 1)) Then
                    p.Location = New Point(content4.LastCartPoint().X, content4.LastCartPoint().Y)
                    content4.pushDeck(stackActual.Cartas(stackActual.Actual - 1))
                    stackActual.popDeck()
                    If stackActual.LastCartPoint <> deck2.LastCartPoint And stackActual.LastCartPoint <> content1.LastCartPoint And stackActual.LastCartPoint <> content2.LastCartPoint And stackActual.LastCartPoint <> content3.LastCartPoint Then
                        stackActual.LastCartPoint = New Point(stackActual.LastCartPoint.X, stackActual.LastCartPoint.Y - 20)
                    End If
                    validarGane()
                    verificarLastPosDeck()
                Else
                    sender.Location = stackActual.LastCartPoint()
                End If
                'verifica que el picture no sea el mismo y este en el rango de la ultima carta disponible de los 7 decks y los 4 stacks de tipos

            ElseIf stackActual.LastCartPoint <> stack1.LastCartPoint And p.Location.Y >= stack1.LastCartPoint.Y - 30 And p.Location.Y <= stack1.LastCartPoint.Y + 150 And
            p.Location.X >= stack1.LastCartPoint.X - 20 And p.Location.X <= stack1.LastCartPoint.X + 70 And stack1.requestDiferentColorAdd(stackActual.Cartas(stackActual.Actual - 1)) And stack1.requestLessNum(stackActual.Cartas(stackActual.Actual - 1)) Then
                p.Location = New Point(stack1.LastCartPoint().X, stack1.LastCartPoint().Y + 20)
                    stack1.LastCartPoint() = New Point(stack1.LastCartPoint().X, stack1.LastCartPoint().Y + 20)
                    stack1.pushDeck(stackActual.Cartas(stackActual.Actual - 1))
                    stackActual.popDeck()
                    If stackActual.LastCartPoint <> deck2.LastCartPoint And stackActual.LastCartPoint <> content1.LastCartPoint And stackActual.LastCartPoint <> content2.LastCartPoint And stackActual.LastCartPoint <> content4.LastCartPoint And stackActual.LastCartPoint <> content3.LastCartPoint Then
                        stackActual.LastCartPoint = New Point(stackActual.LastCartPoint.X, stackActual.LastCartPoint.Y - 20)
                    End If
                    verificarLastPosDeck()

            ElseIf stackActual.LastCartPoint <> stack2.LastCartPoint And p.Location.Y >= stack2.LastCartPoint.Y - 30 And p.Location.Y <= stack2.LastCartPoint.Y + 150 And
            p.Location.X >= stack2.LastCartPoint.X - 20 And p.Location.X <= stack2.LastCartPoint.X + 70 And stack2.requestDiferentColorAdd(stackActual.Cartas(stackActual.Actual - 1)) And stack2.requestLessNum(stackActual.Cartas(stackActual.Actual - 1)) Then
                p.Location = New Point(stack2.LastCartPoint().X, stack2.LastCartPoint().Y + 20)
                stack2.LastCartPoint() = New Point(stack2.LastCartPoint().X, stack2.LastCartPoint().Y + 20)
                stack2.pushDeck(stackActual.Cartas(stackActual.Actual - 1))
                stackActual.popDeck()
                If stackActual.LastCartPoint <> deck2.LastCartPoint And stackActual.LastCartPoint <> content1.LastCartPoint And stackActual.LastCartPoint <> content2.LastCartPoint And stackActual.LastCartPoint <> content4.LastCartPoint And stackActual.LastCartPoint <> content3.LastCartPoint Then
                    stackActual.LastCartPoint = New Point(stackActual.LastCartPoint.X, stackActual.LastCartPoint.Y - 20)
                End If
                verificarLastPosDeck()
            ElseIf stackActual.LastCartPoint <> stack3.LastCartPoint And p.Location.Y >= stack3.LastCartPoint.Y - 30 And p.Location.Y <= stack3.LastCartPoint.Y + 150 And
            p.Location.X >= stack3.LastCartPoint.X - 20 And p.Location.X <= stack3.LastCartPoint.X + 70 And stack3.requestDiferentColorAdd(stackActual.Cartas(stackActual.Actual - 1)) And stack3.requestLessNum(stackActual.Cartas(stackActual.Actual - 1)) Then
                p.Location = New Point(stack3.LastCartPoint().X, stack3.LastCartPoint().Y + 20)
                stack3.LastCartPoint() = New Point(stack3.LastCartPoint().X, stack3.LastCartPoint().Y + 20)
                stack3.pushDeck(stackActual.Cartas(stackActual.Actual - 1))
                stackActual.popDeck()
                If stackActual.LastCartPoint <> deck2.LastCartPoint Then
                    stackActual.LastCartPoint = New Point(stackActual.LastCartPoint.X, stackActual.LastCartPoint.Y - 20)
                End If
                verificarLastPosDeck()
            ElseIf stackActual.LastCartPoint <> stack4.LastCartPoint And p.Location.Y >= stack4.LastCartPoint.Y - 30 And p.Location.Y <= stack4.LastCartPoint.Y + 150 And
            p.Location.X >= stack4.LastCartPoint.X - 20 And p.Location.X <= stack4.LastCartPoint.X + 70 And stack4.requestDiferentColorAdd(stackActual.Cartas(stackActual.Actual - 1)) And stack4.requestLessNum(stackActual.Cartas(stackActual.Actual - 1)) Then
                p.Location = New Point(stack4.LastCartPoint().X, stack4.LastCartPoint().Y + 20)
                stack4.LastCartPoint() = New Point(stack4.LastCartPoint().X, stack4.LastCartPoint().Y + 20)
                stack4.pushDeck(stackActual.Cartas(stackActual.Actual - 1))
                stackActual.popDeck()
                If stackActual.LastCartPoint <> deck2.LastCartPoint And stackActual.LastCartPoint <> content1.LastCartPoint And stackActual.LastCartPoint <> content2.LastCartPoint And stackActual.LastCartPoint <> content4.LastCartPoint And stackActual.LastCartPoint <> content3.LastCartPoint Then
                    stackActual.LastCartPoint = New Point(stackActual.LastCartPoint.X, stackActual.LastCartPoint.Y - 20)
                End If
                verificarLastPosDeck()
            ElseIf stackActual.LastCartPoint <> stack5.LastCartPoint And p.Location.Y >= stack5.LastCartPoint.Y - 30 And p.Location.Y <= stack5.LastCartPoint.Y + 150 And
            p.Location.X >= stack5.LastCartPoint.X - 20 And p.Location.X <= stack5.LastCartPoint.X + 70 And stack5.requestDiferentColorAdd(stackActual.Cartas(stackActual.Actual - 1)) And stack5.requestLessNum(stackActual.Cartas(stackActual.Actual - 1)) Then
                p.Location = New Point(stack5.LastCartPoint().X, stack5.LastCartPoint().Y + 20)
                stack5.LastCartPoint() = New Point(stack5.LastCartPoint().X, stack5.LastCartPoint().Y + 20)
                stack5.pushDeck(stackActual.Cartas(stackActual.Actual - 1))
                stackActual.popDeck()
                If stackActual.LastCartPoint <> deck2.LastCartPoint And stackActual.LastCartPoint <> content1.LastCartPoint And stackActual.LastCartPoint <> content2.LastCartPoint And stackActual.LastCartPoint <> content4.LastCartPoint And stackActual.LastCartPoint <> content3.LastCartPoint Then
                    stackActual.LastCartPoint = New Point(stackActual.LastCartPoint.X, stackActual.LastCartPoint.Y - 20)
                End If
                verificarLastPosDeck()
            ElseIf stackActual.LastCartPoint <> stack6.LastCartPoint And p.Location.Y >= stack6.LastCartPoint.Y - 30 And p.Location.Y <= stack6.LastCartPoint.Y + 150 And
                p.Location.X >= stack6.LastCartPoint.X - 20 And p.Location.X <= stack6.LastCartPoint.X + 70 And stack6.requestDiferentColorAdd(stackActual.Cartas(stackActual.Actual - 1)) And stack6.requestLessNum(stackActual.Cartas(stackActual.Actual - 1)) Then
                p.Location = New Point(stack6.LastCartPoint().X, stack6.LastCartPoint().Y + 20)
                stack6.LastCartPoint() = New Point(stack6.LastCartPoint().X, stack6.LastCartPoint().Y + 20)
                stack6.pushDeck(stackActual.Cartas(stackActual.Actual - 1))
                stackActual.popDeck()
                If stackActual.LastCartPoint <> deck2.LastCartPoint And stackActual.LastCartPoint <> content1.LastCartPoint And stackActual.LastCartPoint <> content2.LastCartPoint And stackActual.LastCartPoint <> content4.LastCartPoint And stackActual.LastCartPoint <> content3.LastCartPoint Then
                    stackActual.LastCartPoint = New Point(stackActual.LastCartPoint.X, stackActual.LastCartPoint.Y - 20)
                End If
                verificarLastPosDeck()
            ElseIf stackActual.LastCartPoint <> stack7.LastCartPoint And p.Location.Y >= stack7.LastCartPoint.Y - 30 And p.Location.Y <= stack7.LastCartPoint.Y + 150 And
                p.Location.X >= stack7.LastCartPoint.X - 20 And p.Location.X <= stack7.LastCartPoint.X + 70 And stack7.requestDiferentColorAdd(stackActual.Cartas(stackActual.Actual - 1)) And stack7.requestLessNum(stackActual.Cartas(stackActual.Actual - 1)) Then
                p.Location = New Point(stack7.LastCartPoint().X, stack7.LastCartPoint().Y + 20)
                stack7.LastCartPoint() = New Point(stack7.LastCartPoint().X, stack7.LastCartPoint().Y + 20)
                stack7.pushDeck(stackActual.Cartas(stackActual.Actual - 1))
                stackActual.popDeck()
                If stackActual.LastCartPoint <> deck2.LastCartPoint And stackActual.LastCartPoint <> content1.LastCartPoint And stackActual.LastCartPoint <> content2.LastCartPoint And stackActual.LastCartPoint <> content4.LastCartPoint And stackActual.LastCartPoint <> content3.LastCartPoint Then
                    stackActual.LastCartPoint = New Point(stackActual.LastCartPoint.X, stackActual.LastCartPoint.Y - 20)
                End If
                verificarLastPosDeck()
            Else
                sender.Location = stackActual.LastCartPoint()
            End If
        Else

            Dim p = DirectCast(sender, PictureBox)
            If stackActual.LastCartPoint <> stack1.LastCartPoint And p.Location.Y >= stack1.LastCartPoint.Y - 30 And p.Location.Y <= stack1.LastCartPoint.Y + 150 And
            p.Location.X >= stack1.LastCartPoint.X - 20 And p.Location.X <= stack1.LastCartPoint.X + 70 And stack1.requestDiferentColorAdd(cartaComparativa) And stack1.requestLessNum(cartaComparativa) Then
                loadNewCarts(stack1)
            ElseIf stackActual.LastCartPoint <> stack2.LastCartPoint And p.Location.Y >= stack2.LastCartPoint.Y - 30 And p.Location.Y <= stack2.LastCartPoint.Y + 150 And
            p.Location.X >= stack2.LastCartPoint.X - 20 And p.Location.X <= stack2.LastCartPoint.X + 70 And stack2.requestDiferentColorAdd(cartaComparativa) And stack2.requestLessNum(cartaComparativa) Then
                loadNewCarts(stack2)
            ElseIf stackActual.LastCartPoint <> stack3.LastCartPoint And p.Location.Y >= stack3.LastCartPoint.Y - 30 And p.Location.Y <= stack3.LastCartPoint.Y + 150 And
            p.Location.X >= stack3.LastCartPoint.X - 20 And p.Location.X <= stack3.LastCartPoint.X + 70 And stack3.requestDiferentColorAdd(cartaComparativa) And stack3.requestLessNum(cartaComparativa) Then
                loadNewCarts(stack3)
            ElseIf stackActual.LastCartPoint <> stack4.LastCartPoint And p.Location.Y >= stack4.LastCartPoint.Y - 30 And p.Location.Y <= stack4.LastCartPoint.Y + 150 And
            p.Location.X >= stack4.LastCartPoint.X - 20 And p.Location.X <= stack4.LastCartPoint.X + 70 And stack4.requestDiferentColorAdd(cartaComparativa) And stack4.requestLessNum(cartaComparativa) Then
                loadNewCarts(stack4)
            ElseIf stackActual.LastCartPoint <> stack5.LastCartPoint And p.Location.Y >= stack5.LastCartPoint.Y - 30 And p.Location.Y <= stack5.LastCartPoint.Y + 150 And
            p.Location.X >= stack5.LastCartPoint.X - 20 And p.Location.X <= stack5.LastCartPoint.X + 70 And stack5.requestDiferentColorAdd(cartaComparativa) And stack5.requestLessNum(cartaComparativa) Then
                loadNewCarts(stack5)
            ElseIf stackActual.LastCartPoint <> stack6.LastCartPoint And p.Location.Y >= stack6.LastCartPoint.Y - 30 And p.Location.Y <= stack6.LastCartPoint.Y + 150 And
            p.Location.X >= stack6.LastCartPoint.X - 20 And p.Location.X <= stack6.LastCartPoint.X + 70 And stack6.requestDiferentColorAdd(cartaComparativa) And stack6.requestLessNum(cartaComparativa) Then
                loadNewCarts(stack6)
            ElseIf stackActual.LastCartPoint <> stack7.LastCartPoint And p.Location.Y >= stack7.LastCartPoint.Y - 30 And p.Location.Y <= stack7.LastCartPoint.Y + 150 And
            p.Location.X >= stack7.LastCartPoint.X - 20 And p.Location.X <= stack7.LastCartPoint.X + 70 And stack7.requestDiferentColorAdd(cartaComparativa) And stack7.requestLessNum(cartaComparativa) Then
                loadNewCarts(stack7)
            Else
                For n = 0 To (listCartsRemoAdd.Count - 1)
                    Dim pic As PictureBox = createCart(listCartsRemoAdd(listCartsRemoAdd.Count - 1 - n))
                    pic.Location = New Point(stackActual.LastCartPoint.X, stackActual.LastCartPoint.Y)
                    stackActual.LastCartPoint = New Point(stackActual.LastCartPoint.X, stackActual.LastCartPoint.Y + 20)
                    stackActual.pushDeck(listCartsRemoAdd(listCartsRemoAdd.Count - 1 - n))
                    listCartsRemoAdd(listCartsRemoAdd.Count - 1 - n).setPic(pic)
                    Me.Controls.Add(pic)
                Next
                stackActual.LastCartPoint = New Point(stackActual.LastCartPoint.X, stackActual.LastCartPoint.Y - 20)
                backImage(stackActual)
            End If
            Me.Controls.Remove(sender)
        End If
        AxWindowsMediaPlayer1.URL = Application.StartupPath() + "\Cartas\Carta.mp3"
    End Sub

    Public Sub loadNewCarts(Pil As Pila)
        Dim a As Integer = 0
        For n = 0 To (listCartsRemoAdd.Count - 1)
            Dim pic As PictureBox = createCart(listCartsRemoAdd(listCartsRemoAdd.Count - 1 - n))
            pic.Location = New Point(Pil.LastCartPoint.X, Pil.LastCartPoint.Y + 20)
            Pil.pushDeck(listCartsRemoAdd(listCartsRemoAdd.Count - 1 - n))
            listCartsRemoAdd(listCartsRemoAdd.Count - 1 - n).setPic(pic)
            Pil.LastCartPoint = New Point(Pil.LastCartPoint.X, Pil.LastCartPoint.Y + 20)
            Me.Controls.Add(pic)
            a += 1
        Next
        If stackActual.Actual <> 0 Then
            Me.Controls.Remove(stackActual.Cartas(stackActual.Actual - 1).Pic)
            stackActual.Cartas(stackActual.Actual - 1).Pic = createCart(stackActual.Cartas(stackActual.Actual - 1))
            stackActual.Cartas(stackActual.Actual - 1).Pic.Location = New Point(stackActual.LastCartPoint.X, stackActual.LastCartPoint.Y - 20)
            Me.Controls.Add(stackActual.Cartas(stackActual.Actual - 1).Pic)
        End If
        stackActual.LastCartPoint = New Point(stackActual.LastCartPoint.X, stackActual.LastCartPoint.Y - 20)
        backImage(Pil)
        backImage(stackActual)
    End Sub

    'Envía los pictures hacia atras de forma que se vean alineados
    Public Sub backImage(pPil As Pila)
        Dim j As Integer = pPil.Actual - 1
        While j <> -1
            pPil.Cartas(j).Pic.SendToBack()
            j -= 1
        End While
    End Sub
    Public Sub verificarLastPosDeck()
        If stackActual.Actual > 0 Then
            Me.Controls.Remove(stackActual.Cartas(stackActual.Actual - 1).Pic)
            stackActual.Cartas(stackActual.Actual - 1).setPic(createCart(stackActual.Cartas(stackActual.Actual - 1)))
            stackActual.Cartas(stackActual.Actual - 1).Pic.Location = stackActual.LastCartPoint
            stackActual.Cartas(stackActual.Actual - 1).Pic.BringToFront()
            Me.Controls.Add(stackActual.Cartas(stackActual.Actual - 1).Pic)
            Dim j As Integer = stackActual.Actual - 1
            While j <> -1
                stackActual.Cartas(j).Pic.SendToBack()
                j -= 1
            End While
        End If
    End Sub

    Public Sub validarGane()
        If content1.Actual <> 0 And content2.Actual <> 0 And content3.Actual <> 0 And content4.Actual <> 0 Then
            If content1.Cartas(content1.Actual - 1).Numero = 13 And content2.Cartas(content2.Actual - 1).Numero = 13 And content3.Cartas(content3.Actual - 1).Numero = 13 And content4.Cartas(content4.Actual - 1).Numero = 13 Then
                MsgBox("Gano")
            End If
        End If
    End Sub

    'Funcion encargada de mostrar y crear las cartas que estaran disponibles en el deck principal
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Me.AxWindowsMediaPlayer1.URL = Application.StartupPath() + "\Cartas\Carta.mp3"
        If Deck.Actual >= 0 Then
            If Deck.Actual > 0 Then
                Dim CartaTEMP As Carta = Deck.Cartas(Deck.Actual - 1)
                Dim Picture As PictureBox = createCart(CartaTEMP)
                CartaTEMP.setPic(Picture)
                deck2.pushDeck(CartaTEMP)
                Deck.popDeck()
                Picture.Size = New Size(90, 130)
                Picture.SizeMode = PictureBoxSizeMode.StretchImage
                CartaTEMP.Pic.Location = New Point(200, 34)
                Me.Controls.Add(Picture)
                Dim j As Integer = deck2.Actual - 1
                While j <> -1
                    If sum1 < 52 Then
                        deck2.Cartas(j).Pic.SendToBack()
                    ElseIf deck2.Actual > 0 Then
                        MsgBox(CStr(deck2.Actual))
                        deck2.Cartas(j).Pic.SendToBack()
                        Picture.Location = New Point(200, 34)
                        Me.Controls.Add(Picture)
                    End If
                    j -= 1
                End While
            Else
                For i = 0 To deck2.Actual - 1
                    Me.Controls.Remove(deck2.Cartas(deck2.Actual - 1).Pic)
                    Deck.pushDeck(deck2.Cartas(deck2.Actual - 1))
                    deck2.popDeck()
                Next
                deck2.removeAllCart()
            End If
        End If
        deck2.LastCartPoint = New Point(200, 34)
    End Sub

    '#########################################

    ' Define la cantidad de cartas iniciales de los decks
    Public Sub settingBoardBegin()
        stack1.defineBegin(1)
        stack2.defineBegin(2)
        stack3.defineBegin(3)
        stack4.defineBegin(4)
        stack5.defineBegin(5)
        stack6.defineBegin(6)
        stack7.defineBegin(7)
        Deck.defineBegin(52)
        insertAllCartsonDeck()
        For i = 0 To 1151
            shuffleDeck()
            shuffleDeck()
            shuffleDeck()
        Next

        'configAllstacks()

    End Sub

    'Crea todas las 52 cartas y las inserta en el deck 
    Public Sub insertAllCartsonDeck()
        Dim Num As Integer = 1
        Dim Tipes As New List(Of String)
        Tipes.Add("C")
        Tipes.Add("D")
        Tipes.Add("E")
        Tipes.Add("T")
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
        For i = 0 To Deck.Actual - 1
            Dim index As Integer
            Dim CartTemporal As New Carta
            index = CInt(New Random(CInt(Date.Now.Ticks And Integer.MaxValue)).Next Mod 51)
            CartTemporal = Deck.Cartas(i)

            Deck.Cartas(i) = Deck.Cartas(index)
            Deck.Cartas(index) = CartTemporal
        Next
    End Sub
    'Genera las cartas iniciales de los 7 stacks de juego dejando el deck solo con 24, colocando las interfaces y únicamente dejando la ultima
    'carta a disposicion, define la posicion de la última carta de la pila
    Public Sub configAllstacks()
        For i = 0 To stack1.Beginer - 1
            Dim cartTemp As New Carta
            Dim pic As New PictureBox
            cartTemp = Deck.Cartas(Deck.Actual - 1)
            pic = createCart(cartTemp)
            cartTemp.setPic(pic)
            stack1.pushDeck(cartTemp)
            Deck.popDeck()
            pic.Location = New Point(20 + 30, 242)
            Me.Controls.Add(pic)
        Next
        stack1.LastCartPoint = New Point(20 + 30, 242)
        For i = 0 To stack2.Beginer - 1
            Dim cartTemp As New Carta
            cartTemp = Deck.Cartas(Deck.Actual - 1)
            If i = stack2.Beginer - 1 Then
                Dim pic As New PictureBox
                pic = createCart(cartTemp)
                pic.Location = New Point(20 + 180, 242 + 20 * i)
                Me.Controls.Add(pic)
                pic.SizeMode = PictureBoxSizeMode.StretchImage
                pic.Cursor = Cursors.Hand
                cartTemp.setPic(pic)
            Else
                Dim pic As New PictureBox
                pic.Size = New Size(90, 130)
                pic.BorderStyle = BorderStyle.FixedSingle
                pic.Image = Image.FromFile(Application.StartupPath() + "\Cartas\Cu.png")
                pic.Location = New Point(20 + 180, 242 + 20 * i)
                pic.SizeMode = PictureBoxSizeMode.StretchImage
                pic.Cursor = Cursors.Hand
                Me.Controls.Add(pic)
                cartTemp.setPic(pic)
            End If
            stack2.pushDeck(cartTemp)
            Deck.popDeck()
            Dim j As Integer = stack2.Actual - 1
            While j >= 0
                stack2.Cartas(j).Pic.SendToBack()
                j -= 1
            End While
        Next
        stack2.LastCartPoint = New Point(20 + 180, 242 + 20 * (stack2.Beginer - 1))
        For i = 0 To stack3.Beginer - 1
            Dim cartTemp As New Carta
            cartTemp = Deck.Cartas(Deck.Actual - 1)
            If i = stack3.Beginer - 1 Then
                Dim pic As New PictureBox
                pic = createCart(cartTemp)
                pic.Location = New Point(20 + 330, 242 + 20 * i)
                Me.Controls.Add(pic)
                pic.SizeMode = PictureBoxSizeMode.StretchImage
                pic.Cursor = Cursors.Hand
                cartTemp.setPic(pic)
            Else
                Dim pic As New PictureBox
                pic.Size = New Size(90, 130)
                pic.BorderStyle = BorderStyle.FixedSingle
                pic.Image = Image.FromFile(Application.StartupPath() + "\Cartas\Cu.png")
                pic.Location = New Point(20 + 330, 242 + 20 * i)
                pic.SizeMode = PictureBoxSizeMode.StretchImage
                pic.Cursor = Cursors.Hand
                Me.Controls.Add(pic)
                cartTemp.setPic(pic)
            End If

            stack3.pushDeck(cartTemp)
            Deck.popDeck()
            Dim j As Integer = stack3.Actual - 1
            While j >= 0
                stack3.Cartas(j).Pic.SendToBack()
                j -= 1
            End While
        Next
        stack3.LastCartPoint = New Point(20 + 330, 242 + 20 * (stack3.Beginer - 1))
        For i = 0 To stack4.Beginer - 1
            Dim cartTemp As New Carta
            cartTemp = Deck.Cartas(Deck.Actual - 1)
            If i = stack4.Beginer - 1 Then
                Dim pic As New PictureBox
                pic = createCart(cartTemp)
                pic.Location = New Point(20 + 480, 242 + 20 * i)
                Me.Controls.Add(pic)
                pic.SizeMode = PictureBoxSizeMode.StretchImage
                pic.Cursor = Cursors.Hand
                cartTemp.setPic(pic)
            Else
                Dim pic As New PictureBox
                pic.Size = New Size(90, 130)
                pic.BorderStyle = BorderStyle.FixedSingle
                pic.Image = Image.FromFile(Application.StartupPath() + "\Cartas\Cu.png")
                pic.Location = New Point(20 + 480, 242 + 20 * i)
                pic.SizeMode = PictureBoxSizeMode.StretchImage
                pic.Cursor = Cursors.Hand
                Me.Controls.Add(pic)
                cartTemp.setPic(pic)
            End If
            stack4.pushDeck(cartTemp)
            Deck.popDeck()
            Dim j As Integer = stack4.Actual - 1
            While j >= 0
                stack4.Cartas(j).Pic.SendToBack()
                j -= 1
            End While
        Next
        stack4.LastCartPoint = New Point(20 + 480, 242 + 20 * (stack4.Beginer - 1))
        For i = 0 To stack5.Beginer - 1
            Dim cartTemp As New Carta
            cartTemp = Deck.Cartas(Deck.Actual - 1)
            If i = stack5.Beginer - 1 Then
                Dim pic As New PictureBox
                pic = createCart(cartTemp)
                pic.Location = New Point(20 + 630, 242 + 20 * i)
                Me.Controls.Add(pic)
                pic.SizeMode = PictureBoxSizeMode.StretchImage
                pic.Cursor = Cursors.Hand
                cartTemp.setPic(pic)
            Else
                Dim pic As New PictureBox
                pic.Size = New Size(90, 130)
                pic.BorderStyle = BorderStyle.FixedSingle
                pic.Image = Image.FromFile(Application.StartupPath() + "\Cartas\Cu.png")
                pic.Location = New Point(20 + 630, 242 + 20 * i)
                pic.SizeMode = PictureBoxSizeMode.StretchImage
                pic.Cursor = Cursors.Hand
                Me.Controls.Add(pic)
                cartTemp.setPic(pic)
            End If
            stack5.pushDeck(cartTemp)
            Deck.popDeck()
            Dim j As Integer = stack5.Actual - 1
            While j >= 0
                stack5.Cartas(j).Pic.SendToBack()
                j -= 1
            End While
        Next
        stack5.LastCartPoint = New Point(20 + 630, 242 + 20 * (stack5.Beginer - 1))
        For i = 0 To stack6.Beginer - 1
            Dim cartTemp As New Carta
            cartTemp = Deck.Cartas(Deck.Actual - 1)
            If i = stack6.Beginer - 1 Then
                Dim pic As New PictureBox
                pic = createCart(cartTemp)
                pic.Location = New Point(20 + 780, 242 + 20 * i)
                Me.Controls.Add(pic)
                pic.SizeMode = PictureBoxSizeMode.StretchImage
                pic.Cursor = Cursors.Hand
                cartTemp.setPic(pic)
            Else
                Dim pic As New PictureBox
                pic.Size = New Size(90, 130)
                pic.BorderStyle = BorderStyle.FixedSingle
                pic.Image = Image.FromFile(Application.StartupPath() + "\Cartas\Cu.png")
                pic.Location = New Point(20 + 780, 242 + 20 * i)
                pic.SizeMode = PictureBoxSizeMode.StretchImage
                pic.Cursor = Cursors.Hand
                Me.Controls.Add(pic)
                cartTemp.setPic(pic)
            End If
            stack6.pushDeck(cartTemp)
            Deck.popDeck()
            Dim j As Integer = stack6.Actual - 1
            While j >= 0
                stack6.Cartas(j).Pic.SendToBack()
                j -= 1
            End While
        Next
        stack6.LastCartPoint = New Point(20 + 780, 242 + 20 * (stack6.Beginer - 1))
        For i = 0 To stack7.Beginer - 1
            Dim cartTemp As New Carta
            cartTemp = Deck.Cartas(Deck.Actual - 1)
            If i = stack7.Beginer - 1 Then
                Dim pic As New PictureBox
                pic = createCart(cartTemp)
                pic.Location = New Point(20 + 930, 242 + 20 * i)
                Me.Controls.Add(pic)
                pic.SizeMode = PictureBoxSizeMode.StretchImage
                pic.Cursor = Cursors.Hand
                cartTemp.setPic(pic)
            Else
                Dim pic As New PictureBox
                pic.Size = New Size(90, 130)
                pic.BorderStyle = BorderStyle.FixedSingle
                pic.Image = Image.FromFile(Application.StartupPath() + "\Cartas\Cu.png")
                pic.Location = New Point(20 + 930, 242 + 20 * i)
                pic.SizeMode = PictureBoxSizeMode.StretchImage
                pic.Cursor = Cursors.Hand
                Me.Controls.Add(pic)
                cartTemp.setPic(pic)
            End If
            stack7.pushDeck(cartTemp)
            Deck.popDeck()
            Dim j As Integer = stack7.Actual - 1
            While j >= 0
                stack7.Cartas(j).Pic.SendToBack()
                j -= 1
            End While
        Next
        stack7.LastCartPoint = New Point(20 + 930, 242 + 20 * (stack7.Beginer - 1))
        content1.LastCartPoint = New Point(500, 34)
        content2.LastCartPoint = New Point(650, 34)
        content3.LastCartPoint = New Point(800, 34)
        content4.LastCartPoint = New Point(950, 34)
    End Sub

    'Crea un picture box a partir de la carta enviada
    Public Function createCart(pCart As Carta) As PictureBox
        Dim pic As New PictureBox
        pic.Size = New Size(90, 130)
        pic.BorderStyle = BorderStyle.FixedSingle
        pic.Image = Image.FromFile(pCart.Imagen)
        pic.SizeMode = PictureBoxSizeMode.StretchImage
        pic.Cursor = Cursors.Hand
        AddHandler pic.MouseUp, AddressOf PictureBox1_MouseUp
        AddHandler pic.MouseDown, AddressOf PictureBox1_MouseDown
        AddHandler pic.MouseMove, AddressOf PictureBox1_MouseMove
        Return pic
    End Function

End Class
