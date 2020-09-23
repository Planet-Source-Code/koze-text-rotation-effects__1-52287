Attribute VB_Name = "TextFunctions"

Public Sub wavetext(pic As PictureBox, Xcord As Long, Ycord As Long, length As Long, curves As Long, height As Long, text As String)
pic.Cls
Dim j
For j = 1 To Len(text)
pic.CurrentX = Xcord + j * length
pic.CurrentY = Ycord - height * Cos(j * (((2 * curves / 360) * 3.14) / Len(text)))
pic.Print Mid(text, j, 1)
Next j
End Sub

Public Sub rotateText(pic As PictureBox, Xcord As Long, Ycord As Long, rad As Long, length As Long, text As String)
pic.Cls
Dim j
For j = 1 To Len(text)
pic.CurrentX = Xcord + length * j * Cos(rad * 3.14 / 180)
pic.CurrentY = Ycord + length * j * Sin(rad * 3.14 / 180)
pic.Print Mid(text, j, 1)
Next j
End Sub
Public Sub circletext(pic As PictureBox, Xcord As Long, Ycord As Long, rad As Long, deg As Long, arc As Long, text As String)
pic.Cls
Dim j
For j = 1 To Len(text)
pic.CurrentX = Xcord + (rad * Sin(j * (((2 * arc / 360) * 3.14) / Len(text)) + 2 * deg * 3.14 / 180))
pic.CurrentY = Ycord - (rad * Cos(j * (((2 * arc / 360) * 3.14) / Len(text)) + 2 * deg * 3.14 / 180))
pic.Print Mid(text, j, 1)
Next j
End Sub
Public Sub eliptictext(pic As PictureBox, Xcord As Long, Ycord As Long, rad As Long, deg As Long, Afactor As Long, Bfactor As Long, arc As Long, text As String)
pic.Cls
Dim j
For j = 1 To Len(text)
pic.CurrentX = Xcord + (rad * Sin(j * (((2 * arc / 360) * 3.14) / Len(text)) + 2 * deg * 3.14 / 180)) / Sqr(Afactor)
pic.CurrentY = Ycord - (rad * Cos(j * (((2 * arc / 360) * 3.14) / Len(text)) + 2 * deg * 3.14 / 180)) / Sqr(Bfactor)
pic.Print Mid(text, j, 1)
Next j
End Sub
