Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim OFD1 As New OpenFileDialog()
        OFD1.Filter = "Open File (*.jpg) | *.jpg | All Files | *.*"
        If OFD1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            PictureBox1.Image = Image.FromFile(OFD1.FileName)
        End If
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim OFD2 As New OpenFileDialog()
        OFD2.Filter = "Open File (*.jpg) | *.jpg | All Files | *.*"
        If OFD2.ShowDialog() = Windows.Forms.DialogResult.OK Then
            PictureBox2.Image = Image.FromFile(OFD2.FileName)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim clr As Integer
        Dim ymax As Integer
        Dim xmax As Integer
        Dim x As Integer
        Dim y As Integer
        Dim bm As Bitmap = New Bitmap(PictureBox1.Image)
        xmax = bm.Width - 1
        ymax = bm.Height - 1
        For y = 0 To ymax
            For x = 0 To xmax
                With bm.GetPixel(x, y)
                    clr = 0.21 * .R + 0.72 * .G + 0.07 * .B
                End With
                bm.SetPixel(x, y, _
                Color.FromArgb(255, clr, clr, clr))
            Next x
        Next y
        PictureBox3.Image = bm
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim clr As Integer
        Dim ymax As Integer
        Dim xmax As Integer
        Dim x As Integer
        Dim y As Integer
        Dim bm2 As Bitmap = New Bitmap(PictureBox1.Image)
        xmax = bm2.Width - 1
        ymax = bm2.Height - 1
        For y = 0 To ymax
            For x = 0 To xmax
                With bm2.GetPixel(x, y)
                    clr = (1 * .R + 1 * .G + 1 * .B) / 3
                End With
                bm2.SetPixel(x, y, _
                Color.FromArgb(255, clr, clr, clr))
            Next x
        Next y
        PictureBox3.Image = bm2
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim clrR As Integer
        Dim clrG As Integer
        Dim clrB As Integer
        Dim ymax As Integer
        Dim xmax As Integer
        Dim x As Integer
        Dim y As Integer
        Dim bm3 As Bitmap = New Bitmap(PictureBox1.Image)
        xmax = bm3.Width - 1
        ymax = bm3.Height - 1
        For y = 0 To ymax
            For x = 0 To xmax
                With bm3.GetPixel(x, y)
                    clrR = Math.Abs((1 * .R) - 255)
                    clrG = Math.Abs((1 * .G) - 255)
                    clrB = Math.Abs((1 * .B) - 255)
                End With
                bm3.SetPixel(x, y, _
                Color.FromArgb(255, clrR, clrG, clrB))
            Next x
        Next y
        PictureBox3.Image = bm3
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim p As Integer
        p = Convert.ToInt32(TextBox1.Text)
        Dim clr As Integer
        Dim ymax As Integer
        Dim xmax As Integer
        Dim x As Integer
        Dim y As Integer
        Dim bm4 As Bitmap = New Bitmap(PictureBox1.Image)
        xmax = bm4.Width - 1
        ymax = bm4.Height - 1
        For y = 0 To ymax
            For x = 0 To xmax
                With bm4.GetPixel(x, y)
                    clr = (1 * .R + 1 * .G + 1 * .B) / 3
                    If clr <= p Then
                        clr = 0
                    End If
                    If clr > p Then
                        clr = 255
                    End If
                End With
                bm4.SetPixel(x, y, _
                Color.FromArgb(255, clr, clr, clr))
            Next x
        Next y
        PictureBox3.Image = bm4
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim p1 As Integer
        Dim p2 As Integer
        p1 = Convert.ToInt32(TextBox5.Text)
        p2 = Convert.ToInt32(TextBox2.Text)
        Dim clr As Integer
        Dim ymax As Integer
        Dim xmax As Integer
        Dim x As Integer
        Dim y As Integer
        Dim bm4 As Bitmap = New Bitmap(PictureBox1.Image)
        xmax = bm4.Width - 1
        ymax = bm4.Height - 1
        For y = 0 To ymax
            For x = 0 To xmax
                With bm4.GetPixel(x, y)
                    clr = (1 * .R + 1 * .G + 1 * .B) / 3
                    If clr < p1 Or clr > p2 Then
                        clr = 0
                    End If
                End With
                bm4.SetPixel(x, y, _
Color.FromArgb(255, clr, clr, clr))
            Next x
        Next y
        PictureBox3.Image = bm4
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim p1 As Integer
        Dim p2 As Integer
        p1 = Convert.ToInt32(TextBox3.Text)
        p2 = Convert.ToInt32(TextBox4.Text)
        Dim clr As Integer
        Dim ymax As Integer
        Dim xmax As Integer
        Dim x As Integer
        Dim y As Integer
        Dim bm5 As Bitmap = New Bitmap(PictureBox1.Image)
        xmax = bm5.Width - 1
        ymax = bm5.Height - 1
        For y = 0 To ymax
            For x = 0 To xmax
                With bm5.GetPixel(x, y)
                    clr = (1 * .R + 1 * .G + 1 * .B) / 3
                    If clr > p1 And clr < p2 Then
                        clr = 0
                        bm5.SetPixel(x, y, _
Color.FromArgb(255, clr, clr, clr))
                    End If
                End With
            Next x
        Next y
        PictureBox3.Image = bm5
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox2_TextChanged_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox2_TextChanged_2(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim p1 As Integer
        Dim p2 As Integer
        p1 = Convert.ToInt32(TextBox7.Text)
        p2 = Convert.ToInt32(TextBox6.Text)
        Dim clr As Integer
        Dim ymax As Integer
        Dim xmax As Integer
        Dim x As Integer
        Dim y As Integer
        Dim bm4 As Bitmap = New Bitmap(PictureBox1.Image)
        xmax = bm4.Width - 1
        ymax = bm4.Height - 1
        For y = 0 To ymax
            For x = 0 To xmax
                With bm4.GetPixel(x, y)
                    clr = (1 * .R + 1 * .G + 1 * .B) / 3
                    If clr < p1 Or clr > p2 Then
                        clr = 255
                    End If
                End With
                bm4.SetPixel(x, y, _
Color.FromArgb(255, clr, clr, clr))
            Next x
        Next y
        PictureBox3.Image = bm4
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim clr As Integer
        Dim clrR As Integer
        Dim clrG As Integer
        Dim clrB As Integer
        Dim ymax As Integer
        Dim xmax As Integer
        Dim x As Integer
        Dim y As Integer
        Dim cont As Integer
        Dim min As Integer
        Dim max As Integer
        Dim bm2 As Bitmap = New Bitmap(PictureBox1.Image)
        xmax = bm2.Width - 1
        ymax = bm2.Height - 1
        For y = 0 To ymax
            For x = 0 To xmax
                With bm2.GetPixel(x, y)
                    clrR = (1 * .R)
                    clrG = (1 * .G)
                    clrB = (1 * .B)
                End With
                min = 255
                max = 0
                If clrR < min Then
                    min = clrR
                End If
                If clrG < min Then
                    min = clrG
                End If
                If clrB < min Then
                    min = clrB
                End If
                If clrR > max Then
                    max = clrR
                End If
                If clrG > max Then
                    max = clrG
                End If
                If clrB > max Then
                    max = clrB
                End If
                clr = (max + min) / 2
                bm2.SetPixel(x, y, _
                Color.FromArgb(255, clr, clr, clr))
            Next x
        Next y
        PictureBox3.Image = bm2
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim p1 As Integer
        Dim p2 As Integer
        Dim clrR As Integer
        Dim clrG As Integer
        Dim clrB As Integer
        Dim clr2R As Integer
        Dim clr2G As Integer
        Dim clr2B As Integer
        Dim clr3R As Integer
        Dim clr3G As Integer
        Dim clr3B As Integer
        Dim ymax As Integer
        Dim xmax As Integer
        Dim x As Integer
        Dim y As Integer
        Dim bm As Bitmap = New Bitmap(PictureBox1.Image)
        Dim bm2 As Bitmap = New Bitmap(PictureBox2.Image)
        Dim bm3 As Bitmap = New Bitmap(PictureBox2.Image)
        xmax = bm.Width - 1
        ymax = bm.Height - 1
        For y = 0 To ymax
            For x = 0 To xmax
                With bm.GetPixel(x, y)
                    clrR = (1 * .R)
                    clrG = (1 * .G)
                    clrB = (1 * .B)
                End With
                With bm2.GetPixel(x, y)
                    clr2R = (1 * .R)
                    clr2G = (1 * .G)
                    clr2B = (1 * .B)
                End With
                clr3R = (clr2R + clrR) / 2
                clr3G = (clr2G + clrG) / 2
                clr3B = (clr2B + clrB) / 2
                bm3.SetPixel(x, y, _
    Color.FromArgb(255, clr3R, clr3G, clr3B))
            Next x
        Next y
        PictureBox3.Image = bm3
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim p1 As Integer
        Dim p2 As Integer
        Dim clrR As Integer
        Dim clrG As Integer
        Dim clrB As Integer
        Dim clr2R As Integer
        Dim clr2G As Integer
        Dim clr2B As Integer
        Dim clr3R As Integer
        Dim clr3G As Integer
        Dim clr3B As Integer
        Dim ymax As Integer
        Dim xmax As Integer
        Dim x As Integer
        Dim y As Integer
        Dim bm As Bitmap = New Bitmap(PictureBox1.Image)
        Dim bm2 As Bitmap = New Bitmap(PictureBox2.Image)
        Dim bm3 As Bitmap = New Bitmap(PictureBox2.Image)
        xmax = bm.Width - 1
        ymax = bm.Height - 1
        For y = 0 To ymax
            For x = 0 To xmax
                With bm.GetPixel(x, y)
                    clrR = (1 * .R)
                    clrG = (1 * .G)
                    clrB = (1 * .B)
                End With
                With bm2.GetPixel(x, y)
                    clr2R = (1 * .R)
                    clr2G = (1 * .G)
                    clr2B = (1 * .B)
                End With
                clr3R = Math.Abs(clrR - clr2R)
                clr3G = Math.Abs(clrG - clr2G)
                clr3B = Math.Abs(clrB - clr2B)
                bm3.SetPixel(x, y, _
    Color.FromArgb(255, clr3R, clr3G, clr3B))
            Next x
        Next y
        PictureBox3.Image = bm3
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dim p1 As Integer
        Dim p2 As Integer
        Dim clrR As Integer
        Dim clrG As Integer
        Dim clrB As Integer
        Dim clr2R As Integer
        Dim clr2G As Integer
        Dim clr2B As Integer
        Dim clr3R As Integer
        Dim clr3G As Integer
        Dim clr3B As Integer
        Dim ymax As Integer
        Dim xmax As Integer
        Dim x As Integer
        Dim y As Integer
        Dim bm As Bitmap = New Bitmap(PictureBox1.Image)
        Dim bm2 As Bitmap = New Bitmap(PictureBox2.Image)
        Dim bm3 As Bitmap = New Bitmap(PictureBox2.Image)
        xmax = bm.Width - 1
        ymax = bm.Height - 1
        For y = 0 To ymax
            For x = 0 To xmax
                With bm.GetPixel(x, y)
                    clrR = (1 * .R)
                    clrG = (1 * .G)
                    clrB = (1 * .B)
                End With
                With bm2.GetPixel(x, y)
                    clr2R = (1 * .R)
                    clr2G = (1 * .G)
                    clr2B = (1 * .B)
                End With
                clr3R = Math.Abs((clrR * clr2R) + clrR)
                clr3G = Math.Abs((clrG * clr2G) + clrG)
                clr3B = Math.Abs((clrB * clr2B) + clrB)
                bm3.SetPixel(x, y, _
    Color.FromArgb(255, clr3R, clr3G, clr3B))
            Next x
        Next y
        PictureBox3.Image = bm3
    End Sub
End Class
