Public Class Form1

    Public fix(14) As Byte
    Public sync(14) As Byte
    Public pos(14) As Point
    Public loc(14) As Point
    Public win(14) As Point
    Public exc(14) As Boolean
    Public clicked(14) As Boolean
    Public trigger As Boolean

    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Byte) As Byte

    Sub GetCurrPos()
        pos(0) = Button13.Location
        pos(1) = Button1.Location
        pos(2) = Button2.Location
        pos(3) = Button3.Location
        pos(4) = Button7.Location
        pos(5) = Button6.Location
        pos(6) = Button5.Location
        pos(7) = Button4.Location
        pos(8) = Button11.Location
        pos(9) = Button10.Location
        pos(10) = Button9.Location
        pos(11) = Button8.Location
        pos(12) = Button19.Location
        pos(13) = Button18.Location
        pos(14) = Button17.Location
    End Sub

    Sub Simulation(button As Byte) ' simulate manipulation position of buttons
        trigger = False
        If GetAsyncKeyState(Keys.S) And loc(button).Y <> 147 Then ' down
            loc(button).Y += 49
        ElseIf GetAsyncKeyState(Keys.D) And loc(button).X <> 228 Then ' right
            loc(button).X += 76
        ElseIf GetAsyncKeyState(Keys.W) And loc(button).Y <> 0 Then ' up
            loc(button).Y -= 49
        ElseIf GetAsyncKeyState(Keys.A) And loc(button).X <> 0 Then ' left
            loc(button).X -= 76
        End If
        GetCurrPos()
        TriggerCheck(button)
        If button = 0 And clicked(0) = True And trigger = False Then ' 1
            Button13.Location = loc(button)
        ElseIf button = 1 And clicked(1) = True And trigger = False Then ' 2
            Button1.Location = loc(button)
        ElseIf button = 2 And clicked(2) = True And trigger = False Then ' 3
            Button2.Location = loc(button)
        ElseIf button = 3 And clicked(3) = True And trigger = False Then ' 4
            Button3.Location = loc(button)
        ElseIf button = 4 And clicked(4) = True And trigger = False Then ' 5
            Button7.Location = loc(button)
        ElseIf button = 5 And clicked(5) = True And trigger = False Then ' 6
            Button6.Location = loc(button)
        ElseIf button = 6 And clicked(6) = True And trigger = False Then ' 7
            Button5.Location = loc(button)
        ElseIf button = 7 And clicked(7) = True And trigger = False Then ' 8
            Button4.Location = loc(button)
        ElseIf button = 8 And clicked(8) = True And trigger = False Then ' 9
            Button11.Location = loc(button)
        ElseIf button = 9 And clicked(9) = True And trigger = False Then ' 10
            Button10.Location = loc(button)
        ElseIf button = 10 And clicked(10) = True And trigger = False Then ' 11
            Button9.Location = loc(button)
        ElseIf button = 11 And clicked(11) = True And trigger = False Then ' 12
            Button8.Location = loc(button)
        ElseIf button = 12 And clicked(12) = True And trigger = False Then ' 13
            Button19.Location = loc(button)
        ElseIf button = 13 And clicked(13) = True And trigger = False Then ' 14
            Button18.Location = loc(button)
        ElseIf button = 14 And clicked(14) = True And trigger = False Then ' 15
            Button17.Location = loc(button)
        End If

        If button = 0 And clicked(0) = True And trigger = True Then
            Button13.Location = pos(button)
            loc(button) = pos(button)
        ElseIf button = 1 And clicked(1) = True And trigger = True Then
            Button1.Location = pos(button)
            loc(button) = pos(button)
        ElseIf button = 2 And clicked(2) = True And trigger = True Then
            Button2.Location = pos(button)
            loc(button) = pos(button)
        ElseIf button = 3 And clicked(3) = True And trigger = True Then
            Button3.Location = pos(button)
            loc(button) = pos(button)
        ElseIf button = 4 And clicked(4) = True And trigger = True Then
            Button7.Location = pos(button)
            loc(button) = pos(button)
        ElseIf button = 5 And clicked(5) = True And trigger = True Then
            Button6.Location = pos(button)
            loc(button) = pos(button)
        ElseIf button = 6 And clicked(6) = True And trigger = True Then
            Button5.Location = pos(button)
            loc(button) = pos(button)
        ElseIf button = 7 And clicked(7) = True And trigger = True Then
            Button4.Location = pos(button)
            loc(button) = pos(button)
        ElseIf button = 8 And clicked(8) = True And trigger = True Then
            Button11.Location = pos(button)
            loc(button) = pos(button)
        ElseIf button = 9 And clicked(9) = True And trigger = True Then
            Button10.Location = pos(button)
            loc(button) = pos(button)
        ElseIf button = 10 And clicked(10) = True And trigger = True Then
            Button9.Location = pos(button)
            loc(button) = pos(button)
        ElseIf button = 11 And clicked(11) = True And trigger = True Then
            Button8.Location = pos(button)
            loc(button) = pos(button)
        ElseIf button = 12 And clicked(12) = True And trigger = True Then
            Button19.Location = pos(button)
            loc(button) = pos(button)
        ElseIf button = 13 And clicked(13) = True And trigger = True Then
            Button18.Location = pos(button)
            loc(button) = pos(button)
        ElseIf button = 14 And clicked(14) = True And trigger = True Then
            Button17.Location = pos(button)
            loc(button) = pos(button)
        End If
    End Sub

    Sub CheckWinPos()
        Dim ended As Boolean
        For i = 0 To 14
            If loc(i) = win(i) Then
                exc(i) = True
            End If
        Next
        If exc(0) = True And exc(1) = True And exc(2) = True And exc(3) = True And exc(4) = True And exc(5) = True And exc(6) = True And exc(7) = True And exc(8) = True And exc(9) = True And exc(10) = True And exc(11) = True And exc(12) = True And exc(13) = True And exc(14) = True Then
            MsgBox("Поздравляем! Вы выиграли!", vbOK + vbInformation)
            ended = True
        End If
        If ended = True Then
            End
        End If
    End Sub

    Private Sub TriggerCheck(button As Byte)
        For i = 0 To 14
            If loc(button) = pos(i) Then
                trigger = True
            End If
        Next
    End Sub

    Sub UpdateGameHandle() ' update data buttons
        If clicked(0) = True Then ' 1
            Simulation(0)
            clicked(0) = False
        ElseIf clicked(1) = True Then ' 2
            Simulation(1)
            clicked(1) = False
        ElseIf clicked(2) = True Then ' 3
            Simulation(2)
            clicked(2) = False
        ElseIf clicked(3) = True Then ' 4
            Simulation(3)
            clicked(3) = False
        ElseIf clicked(4) = True Then ' 5
            Simulation(4)
            clicked(4) = False
        ElseIf clicked(5) = True Then ' 6
            Simulation(5)
            clicked(5) = False
        ElseIf clicked(6) = True Then ' 7
            Simulation(6)
            clicked(6) = False
        ElseIf clicked(7) = True Then ' 8
            Simulation(7)
            clicked(7) = False
        ElseIf clicked(8) = True Then ' 9
            Simulation(8)
            clicked(8) = False
        ElseIf clicked(9) = True Then ' 10
            Simulation(9)
            clicked(9) = False
        ElseIf clicked(10) = True Then ' 11
            Simulation(10)
            clicked(10) = False
        ElseIf clicked(11) = True Then ' 12
            Simulation(11)
            clicked(11) = False
        ElseIf clicked(12) = True Then ' 13
            Simulation(12)
            clicked(12) = False
        ElseIf clicked(13) = True Then ' 14
            Simulation(13)
            clicked(13) = False
        ElseIf clicked(14) = True Then ' 15
            Simulation(14)
            clicked(14) = False
        End If
        CheckWinPos()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        ' 1
        clicked(0) = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '2
        clicked(1) = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '3
        clicked(2) = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        '4
        clicked(3) = True
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        '5
        clicked(4) = True
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        '6
        clicked(5) = True
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        '7
        clicked(6) = True
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        '8
        clicked(7) = True
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        '9
        clicked(8) = True
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        '10
        clicked(9) = True
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        '11
        clicked(10) = True
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        '12
        clicked(11) = True
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        '13
        clicked(12) = True
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        '14
        clicked(13) = True
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        '15
        clicked(14) = True
    End Sub

    Public Sub RandomizeValues() ' omg i fix this shit!
        Dim xRND As Integer

        For i = 0 To 14
            sync(i) = i
        Next

        For xi = 0 To 14 : sync(xi) = xi : Next xi

        For xi = 0 To 14
            xRND = Int(Rnd() * ((14 + 1) - xi))
            fix(xi) = sync(xRND)
            sync(xRND) = sync(14 - xi)
        Next xi
    End Sub

    Sub UpdatePosition(stepA As Byte) '' shitiest method but idk that i must do.... / save start pos all buttons
        loc(0) = Button13.Location
        loc(1) = Button1.Location
        loc(2) = Button2.Location
        loc(3) = Button3.Location
        loc(4) = Button7.Location
        loc(5) = Button6.Location
        loc(6) = Button5.Location
        loc(7) = Button4.Location
        loc(8) = Button11.Location
        loc(9) = Button10.Location
        loc(10) = Button9.Location
        loc(11) = Button8.Location
        loc(12) = Button19.Location
        loc(13) = Button18.Location
        loc(14) = Button17.Location
        If stepA = 1 Then
            For i = 0 To 14
                win(i) = loc(i) 'get all win positions
            Next
        End If
    End Sub

    Sub LoadPosition() ' threading - 1 method fix but values not unikal
        RandomizeValues()
        Button13.Location = loc(fix(0))
        Threading.Thread.Sleep(1)
        Button1.Location = loc(fix(1))
        Threading.Thread.Sleep(1)
        Button2.Location = loc(fix(2))
        Threading.Thread.Sleep(1)
        Button3.Location = loc(fix(3))
        Threading.Thread.Sleep(1)
        Button7.Location = loc(fix(4))
        Threading.Thread.Sleep(1)
        Button6.Location = loc(fix(5))
        Threading.Thread.Sleep(1)
        Button5.Location = loc(fix(6))
        Threading.Thread.Sleep(1)
        Button4.Location = loc(fix(7))
        Threading.Thread.Sleep(1)
        Button11.Location = loc(fix(8))
        Threading.Thread.Sleep(1)
        Button10.Location = loc(fix(9))
        Threading.Thread.Sleep(1)
        Button9.Location = loc(fix(10))
        Threading.Thread.Sleep(1)
        Button8.Location = loc(fix(11))
        Threading.Thread.Sleep(1)
        Button19.Location = loc(fix(12))
        Threading.Thread.Sleep(1)
        Button18.Location = loc(fix(13))
        Threading.Thread.Sleep(1)
        Button17.Location = loc(fix(14))
        GetCurrPos()

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        ' start game
        Button12.Enabled = False
        UpdatePosition(1)
        Threading.Thread.Sleep(100)
        LoadPosition()
        Threading.Thread.Sleep(100)
        UpdatePosition(0)
        Button14.Enabled = True
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        UpdateGameHandle()
    End Sub
End Class
