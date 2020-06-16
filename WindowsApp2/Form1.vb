Public Class Form1
    Dim T As Boolean 'change player
    Dim a As Integer
    Private Sub c1_click(sender As Object, e As EventArgs) Handles c1.Click, c2.Click, c3.Click, c4.Click, c5.Click, c6.Click, c7.Click, c8.Click, c9.Click
        If Label1.Text = "" Then '未分出勝負
            If sender.Text = "" Then '棋格是空的
                If T = True Then
                    sender.Text = "X" '畫X
                    a = a + 1
                    Button1.Enabled = False
                    Button2.Enabled = False
                Else
                    sender.Text = "O" '畫O
                    a = a + 1
                    Button1.Enabled = False
                    Button2.Enabled = False
                End If
                T = Not (T) 'change player
                If chkWin() <> "" Then
                    Label1.Text = chkWin() + "贏了"
                    PictureBox1.Visible = True
                    My.Computer.Audio.Play(My.Application.Info.DirectoryPath & "\歡呼.wav",
                    AudioPlayMode.Background)
                    Button1.Enabled = True
                    Button2.Enabled = True
                ElseIf a = 9 Then
                    Label1.Text = "平手"
                        a = 0
                        Button1.Enabled = True
                    Button2.Enabled = True
                End If

            End If
        End If
    End Sub
    '離開遊戲
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        My.Computer.Audio.Stop()
    End Sub
    '清空棋盤
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load, Button1.Click
        For Each c In Me.Controls  '表單上的每一個controls
            If TypeOf (c) Is Label Then '如果是標籤物件
                c.text = "" '清除文字內容
                PictureBox1.Visible = False
                a = 0
                My.Computer.Audio.Stop()
            End If
        Next
    End Sub
    '判斷勝負
    Function chkWin() As String
        If c1.Text + c2.Text + c3.Text = "OOO" Then Return "O"
        If c4.Text + c5.Text + c6.Text = "OOO" Then Return "O"
        If c7.Text + c8.Text + c9.Text = "OOO" Then Return "O"
        If c1.Text + c4.Text + c7.Text = "OOO" Then Return "O"
        If c2.Text + c5.Text + c8.Text = "OOO" Then Return "O"
        If c3.Text + c6.Text + c9.Text = "OOO" Then Return "O"
        If c1.Text + c5.Text + c9.Text = "OOO" Then Return "O"
        If c3.Text + c5.Text + c7.Text = "OOO" Then Return "O"
        If c1.Text + c2.Text + c3.Text = "XXX" Then Return "X"
        If c4.Text + c5.Text + c6.Text = "XXX" Then Return "X"
        If c7.Text + c8.Text + c9.Text = "XXX" Then Return "X"
        If c1.Text + c4.Text + c7.Text = "XXX" Then Return "X"
        If c2.Text + c5.Text + c8.Text = "XXX" Then Return "X"
        If c3.Text + c6.Text + c9.Text = "XXX" Then Return "X"
        If c1.Text + c5.Text + c9.Text = "XXX" Then Return "X"
        If c3.Text + c5.Text + c7.Text = "XXX" Then Return "X"
        Return "" '圈叉都未連線
    End Function
End Class