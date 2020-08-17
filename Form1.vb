Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim pic() As PictureBox 'create a variable that stores a picturebox array
        pic = New PictureBox() {PictureBox1, PictureBox2, PictureBox3, PictureBox4, PictureBox5}
        pic(0) = PictureBox1
        pic(1) = PictureBox2
        pic(2) = PictureBox3
        pic(3) = PictureBox4
        pic(4) = PictureBox5
        Dim randomNumbers(5) As Integer 'array of random numbers
        Dim randomStars(1) As Integer 'array of random stars
        Dim value As String
        Dim i As Integer
        Dim num As Integer = Val(ComboBox1.Text) 'the variable "num" receives the value of the combobox
        If num > 5 Or num <= 0 Then 'checks if the value entered in the combobox is greater than 5 or less than 0
            MsgBox("INSIRA UM VALOR VALIDO", MsgBoxStyle.Critical) ' send error message 

        ElseIf ListBox1.Items.Count > 0 Then 'check if there is anything in the listbox if it exists it asks to clean
            MsgBox("LIMPE A TELA ANTES DE GERAR UM NOVO CODIGO", MsgBoxStyle.Information) ' send info message 
        Else
            For i = 0 To num - 1 'after checking, loop from 0 to "num"
                Dim orderedKey As String = ""
                Dim orderedKey2 As String = ""
                Dim intArray = GenerateRandomNumber(randomNumbers, 1, 51) 'calls the function passing the parameters of the array size, initial and final value
                Dim stars = GenerateRandomNumber(randomStars, 1, 13) 'calls the function passing the parameters of the array size, initial and final value
                For Each value In intArray
                    orderedKey = orderedKey & " - " & value
                Next
                orderedKey = orderedKey.Substring(3) 'subtracts the amount of strings passed in the substring parameter

                For Each value In stars
                    orderedKey2 = orderedKey2 & " - " & value
                Next
                orderedKey2 = orderedKey2.Substring(3) 'subtracts the amount of strings passed in the substring parameter
                ListBox1.Items.Insert((i), orderedKey) 'inserts from the "i" the "orderedkey" value in the listbox
                ListBox2.Items.Insert((i), orderedKey2) 'inserts from the "i" the "orderedkey2" value in the listbox

                DrawNumbers(intArray, pic(i)) 'calls the function passing the numbers already ordered and in which picturebox you want to draw
                Drawstars(stars, pic(i)) 'calls the function passing the stars already ordered and in which picturebox you want to draw
            Next
        End If

    End Sub
    Public Sub DrawNumbers(ByVal drawnKeys() As Integer, ByVal imageCard As PictureBox)
        Dim myPen As Pen 'creates a variable with a pen attribute
        myPen = New Pen(Drawing.Color.Black, 3) 'in the variable put the color you want and the thickness of the stroke
        Dim myGraphics As Graphics = imageCard.CreateGraphics 'create a variable that becomes the object to be drawn, that object and a picturebox received in the function
        Dim i As Integer
        Dim x As Integer = -22 'value that starts point x to make the circle, left or rigth
        Dim y As Integer = 10 'value that starts point y to make the circle, up or down
        Dim z As Integer = 28 'distance the circle is from the other
        For i = 0 To drawnKeys.Length - 1 'cycle that runs through the array size -1
            Select Case drawnKeys(i)'checks the value one by one, and drawing in the respective coordinate specified in each case
                Case < 7 'Case 1, 2, 3, 4, 5, 6
                    myGraphics.DrawEllipse(myPen, x + (z * drawnKeys(i)), y, 15, 15)
                                           'the pen, x coordinate, y coordinate, length and height to make the perfect circle
                Case 7 To 12 'Case 7, 8, 9, 10, 11, 12
                    myGraphics.DrawEllipse(myPen, x + (z * (drawnKeys(i) - 6)), y + (z * 1), 15, 15)
                Case 13 To 18
                    myGraphics.DrawEllipse(myPen, x + (z * (drawnKeys(i) - 12)), y + (z * 2), 15, 15)
                Case 19 To 24
                    myGraphics.DrawEllipse(myPen, x + (z * (drawnKeys(i) - 18)), y + (z * 3), 15, 15)
                Case 25 To 30
                    myGraphics.DrawEllipse(myPen, x + (z * (drawnKeys(i) - 24)), y + (z * 4), 15, 15)
                Case 31 To 36
                    myGraphics.DrawEllipse(myPen, x + (z * (drawnKeys(i) - 30)), y + (z * 5), 15, 15)
                Case 37 To 42
                    myGraphics.DrawEllipse(myPen, x + (z * (drawnKeys(i) - 36)), y + (z * 6), 15, 15)
                Case 43 To 48
                    myGraphics.DrawEllipse(myPen, x + (z * (drawnKeys(i) - 42)), y + (z * 7), 15, 15)
                Case 49, 50
                    myGraphics.DrawEllipse(myPen, x + (z * (drawnKeys(i) - 48)), y + (z * 8), 15, 15)
                Case Else
            End Select

        Next

    End Sub
    Public Sub Drawstars(ByVal drawnStars() As Integer, ByVal imageCard As PictureBox) 'function to draw the stars, receives as a parameter the stars ordered in the array and the image box to be drawn
        Dim myPen As Pen
        myPen = New Pen(Drawing.Color.Green, 3)
        Dim myGraphics As Graphics = imageCard.CreateGraphics
        Dim i As Integer
        Dim x As Integer = -24
        Dim y As Integer = 275 'value varies according to the position of the image you want to draw
        Dim z As Integer = 47

        For i = 0 To drawnStars.Length - 1
            Select Case drawnStars(i)
                Case < 4
                    myGraphics.DrawEllipse(myPen, x + (z * drawnStars(i)), y, 15, 15)
                Case 4 To 6
                    myGraphics.DrawEllipse(myPen, x + ((z + 3) * (drawnStars(i) - 3)), (y - 6) + (z * 1), 15, 15)
                Case 7 To 9
                    myGraphics.DrawEllipse(myPen, x + ((z + 3) * (drawnStars(i) - 6)), (y - 12) + (z * 2), 15, 15)
                Case 10 To 12
                    myGraphics.DrawEllipse(myPen, x + ((z + 3) * (drawnStars(i) - 9)), (y - 14) + (z * 3), 15, 15)
                Case Else
            End Select

        Next

    End Sub
    Public Function GetRandom(ByVal Min As Integer, ByVal Max As Integer) As Integer
        ' by making Generator static, we preserve the same instance '
        ' (i.e., do not create new instances with the same seed over and over) '
        ' between calls '
        Static Generator As System.Random = New System.Random()
        Return Generator.Next(Min, Max)
    End Function

    Public Function GenerateRandomNumber(ByVal randoms() As Integer, ByVal initial As Integer, ByVal final As Integer)
        Dim unorderedKey As String
        unorderedKey = ""
        Dim number As Integer
        Dim i As Integer
        Dim equalNumber As Boolean

        For i = 0 To randoms.Length - 1 'just go through the array size
            equalNumber = False ' always starts false
            number = GetRandom(initial, final) ' randomly generates a number between initial and final
            For y As Integer = 0 To i
                If (randoms(y) = number) Then 'checks if in position y is equal to the number generated
                    equalNumber = True ' the "equalnumber" becomes true if the numbers has been found in the array
                    Exit For
                End If
            Next
            If (equalNumber) Then ' if the number exists in the array then the "i" returns  
                i = i - 1 ' i-1 to be able to generate the number again
            Else
                randoms(i) = number 'if the number does not exist in then it is inserted
                unorderedKey = unorderedKey & " - " & randoms(i)
            End If

        Next
        Array.Sort(randoms)

        Return randoms

    End Function
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ListBox1.Items.Clear() 'clear the listbox
        ListBox2.Items.Clear()
        PictureBox1.Image = Nothing 'clear the picturebox
        PictureBox2.Image = Nothing
        PictureBox3.Image = Nothing
        PictureBox4.Image = Nothing
        PictureBox5.Image = Nothing
        ComboBox1.Text = Nothing ' clear the comboBox
    End Sub
End Class
