Public Class Form1

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then TextBox1.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then TextBox2.Text = FolderBrowserDialog1.SelectedPath
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If FolderBrowserDialog2.ShowDialog() = DialogResult.OK Then TextBox4.Text = FolderBrowserDialog2.SelectedPath
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If SaveFileDialog1.ShowDialog = DialogResult.OK Then TextBox3.Text = SaveFileDialog1.FileName
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim folder As New zipFolder(TextBox4.Text & "\")
        folder.saveTo(TextBox3.Text)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim zip As New zipFile(TextBox1.Text)
        zip.extractTo(TextBox2.Text & "\")
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        OpenFileDialog1.ShowDialog()
        TextBox5.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim PKHDR As Byte() = {Asc("P"), Asc("K"), 5, 6}
        Dim Zip As Byte()
        Zip = IO.File.ReadAllBytes(TextBox5.Text)

        Dim HeaderLoc As Integer

        For i As Integer = Zip.Length - PKHDR.Length To 0 Step -1
            HeaderLoc = 0
            For J As Integer = 0 To PKHDR.Length - 1
                If Not PKHDR(J) = Zip(i + J) Then HeaderLoc = -1
            Next
            If HeaderLoc = 0 Then
                HeaderLoc = i
                Exit For
            End If
        Next

        ReDim Preserve Zip(TextBox6.Text - 1)

        Dim holder As Integer
        For i As Integer = 0 To 53
            holder = Zip(HeaderLoc + i) 'This will avoid corruption on files with < HEADER_SIZE bytes open
            Zip(HeaderLoc + i) = 0
            Zip(Zip.Length - 54 + i) = holder
        Next
        IO.File.WriteAllBytes(TextBox5.Text, Zip)
    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
