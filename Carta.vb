Public Class Carta
    Public Property Tipo As String
    Public Property Color As String
    Public Property Numero As Integer
    Public Property Imagen As String = Application.StartupPath()

    Public Property Pic As PictureBox
    Public Property Position As Point

    Public Sub defineCart(pTipe As String, pNumero As Integer)
        Tipo = pTipe
        Numero = pNumero
        If pTipe = "C" Or pTipe = "D" Then
            Color = "Rojo"
        ElseIf pTipe = "E" Or pTipe = "T" Then
            Color = "Negro"
        End If
        Imagen += "\Cartas\" + pTipe + CStr(pNumero) + ".png"
    End Sub

    Public Sub setPic(picturebox As PictureBox)
        Pic = picturebox
    End Sub

End Class
