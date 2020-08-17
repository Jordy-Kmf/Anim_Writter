Imports System.Runtime.InteropServices
Public Class Form1
    Dim wait As Integer = 3

    'display backC
    Dim tR As Integer = BackColor.R
    Dim tG As Integer = BackColor.G
    Dim tB As Integer = BackColor.B

    <DllImport("user32.DLL", EntryPoint:="ReleaseCapture")>
    Private Shared Sub ReleaseCapture()
    End Sub

    <DllImport("user32.DLL", EntryPoint:="SendMessage")>
    Private Shared Sub SendMessage(ByVal hWnd As System.IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer)
    End Sub

    Private Sub pnlTitleBar_MouseMove(sender As Object, e As MouseEventArgs) Handles pnlTitleBar.MouseMove
        ReleaseCapture()
        SendMessage(Me.Handle, &H112&, &HF012&, 0)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim fonts As New Drawing.Text.InstalledFontCollection()
        For fntFamily As Integer = 0 To fonts.Families.Length - 1
            cbxFont.Items.Add(fonts.Families(fntFamily).Name)
        Next

        'History callouts
        Dim sC = My.Settings.StartColor
        Dim cToUse = xtractColor(sC)
        Dim bR As Integer = cToUse.r
        Dim bG As Integer = cToUse.g
        Dim bB As Integer = cToUse.b

        BackColor = findBack(bR, bG, bB)
        pnlTitleBar.BackColor = cToUse
        pnlBase.BackColor = cToUse
        btnSideBar.BackColor = cToUse
        btnSideBarR.BackColor = cToUse

        cbxFont.Text = My.Settings.Font
        numFontSize.Value = My.Settings.FontSize
        numAniInt.Value = My.Settings.AnimationSpeed
        numDelay.Value = My.Settings.StartDelay
        bxInput.Text = My.Settings.TextInput
        bxInput.Font = New Font(cbxFont.Text, numFontSize.Value, bxInput.SelectionFont.Style)

        'For before-deployment use ++++++++++++++++++++++++
        'bxDisplay.BackColor = Color.FromArgb(211, 197, 237)
        'bxDisplay.BackColor = cToUse
        dBox()
        'Others
        lblSec.Text = sec & " seconds"

    End Sub

#Region "Functions"
    Function PushOrPull()
        'Open
        If pnlSideBase.Location = New Point(-270, 38) Then
            pnlSideBase.Location = New Point(-270, 38)
            Do Until pnlSideBase.Location.X = 0
                pnlSideBase.Location = New Point(pnlSideBase.Location.X + 10, 38)
                Refresh()
            Loop
            Do Until pnlSideBase.Location.X = 0
                pnlSideBase.Location = New Point(pnlSideBase.Location.X + 10, 38)
                Refresh()
                System.Threading.Thread.Sleep(20)
            Loop
            bxInput.Focus()
        Else 'Close
            pnlSideBase.Location = New Point(0, 38)
            Do Until pnlSideBase.Location.X = -270
                pnlSideBase.Location = New Point(pnlSideBase.Location.X - 10, 38)
                Refresh()
            Loop
            Do Until pnlSideBase.Location.X = -270
                pnlSideBase.Location = New Point(pnlSideBase.Location.X - 10, 38)
                Refresh()
                System.Threading.Thread.Sleep(20)
            Loop
        End If
        Return 0
    End Function

    Function PushOrPullM()
        'Open
        If pnlMusic.Location = New Point(513, 38) Then
            pnlMusic.Location = New Point(513, 38)
            Do Until pnlMusic.Location.X = 283
                pnlMusic.Location = New Point(pnlMusic.Location.X - 10, 38)
                Refresh()
            Loop
            Do Until pnlMusic.Location.X = 283
                pnlMusic.Location = New Point(pnlMusic.Location.X - 10, 38)
                Refresh()
                System.Threading.Thread.Sleep(20)
            Loop
            bxInput.Focus()
            'Close
        Else
            pnlMusic.Location = New Point(283, 38)
            Do Until pnlMusic.Location.X = 513
                pnlMusic.Location = New Point(pnlMusic.Location.X + 10, 38)
                Refresh()
            Loop
            Do Until pnlMusic.Location.X = 513
                pnlMusic.Location = New Point(pnlMusic.Location.X + 10, 38)
                Refresh()
                System.Threading.Thread.Sleep(20)
            Loop
        End If
        Return 0
    End Function

    Function ShowMe()
        btnClose.Visible = True
        btnMinimise.Visible = True
    End Function

    Function HideMe()
        btnClose.Visible = False
        btnMinimise.Visible = False
        tmHide.Enabled = False
    End Function

    Function xtractColor(sC)
        Dim lR As Integer = sC.IndexOf(",")
        Dim cR = sC.Substring(0, lR)
        sC = sC.Substring(lR + 1)
        Dim lG As Integer = sC.IndexOf(",")
        Dim cG = sC.Substring(0, lG)
        Dim cB As Integer = sC.Substring(lG + 1)
        Return Color.FromArgb(255, cR, cG, cB)
    End Function

    Function changeColor(btn)
        Dim btnNN As Button = CType(btn, Button)
        Dim r As Integer = btnNN.BackColor.R
        Dim b As Integer = btnNN.BackColor.B
        Dim g As Integer = btnNN.BackColor.G

        If r = 0 Or r < 220 Then
            r += 35
        ElseIf r = 255 Or r < 255 Then
            r -= 35
        End If

        If g = 0 Or g < 220 Then
            g += 35
        ElseIf g = 255 Or g < 255 Then
            g -= 35
        End If

        Return Color.FromArgb(r, g, b)
    End Function

    Function findBack(colorR, colorG, colorB)
        If colorR = 0 Or colorR < 220 Then
            colorR += 35
        ElseIf colorR = 255 Or colorR < 255 Then
            colorR -= 35
        End If

        If colorG = 0 Or colorG < 220 Then
            colorG += 35
        ElseIf colorG = 255 Or colorG < 255 Then
            colorG -= 35
        End If

        Return Color.FromArgb(colorR, colorG, colorB)
    End Function

    Function dBox()
        If bxDisplay.BackColor.GetBrightness > 0.4 Then
            bxDisplay.ForeColor = Color.FromArgb(64, 64, 64)
        Else bxDisplay.ForeColor = Color.WhiteSmoke
        End If

    End Function

    Function durationCalc(s As Integer, t As Integer)
        s = numAniInt.Value
        t = bxInput.TextLength

        Dim d As Integer = (s / 1000) * t
        'lblDuration.Text = d & " seconds"
        Return d
    End Function

#End Region

    Private Sub btnSideBar_Click(sender As Object, e As EventArgs) Handles btnSideBar.Click
        PushOrPull()
        bxInput.Focus()
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        My.Settings.StartColor = pnlTitleBar.BackColor.R & "," & pnlTitleBar.BackColor.G & "," & pnlTitleBar.BackColor.B
        My.Settings.Font = cbxFont.Text
        My.Settings.FontSize = numFontSize.Value
        My.Settings.AnimationSpeed = numAniInt.Value
        My.Settings.StartDelay = numDelay.Value
        My.Settings.TextInput = bxInput.Text
        My.Settings.Save()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Application.Exit()
    End Sub

    Private Sub btnMinimise_Click(sender As Object, e As EventArgs) Handles btnMinimise.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

#Region "Timers"
    Dim inter As Integer = 1
    Private Sub tmHide_Tick(sender As Object, e As EventArgs) Handles tmHide.Tick
        inter += 1      'Control box delay
        If inter = 3 Then
            inter = 1
            HideMe()
        End If
    End Sub

    Private Sub tmAnimate_Tick(sender As Object, e As EventArgs) Handles tmAnimate.Tick
        If bxDisplay.Text.Length = str.Length Then
            tmDuration.Enabled = False 'May be removed
            tmAnimate.Stop()
            wait = 2
            tmDone.Enabled = True
            'The rest of the action is sent to tmDone()
            Exit Sub
        End If

        bxDisplay.Text = str.Substring(0, count)        'Writer
        count = count + 1
    End Sub

    Private Sub tmDelay_Tick(sender As Object, e As EventArgs) Handles tmDelay.Tick
        If wait = -1 Then        'Ready to go
            lblDelay.Visible = False
            lblTag.ForeColor = lblDelay.ForeColor
            tmDelay.Enabled = False
            bxDisplay.Text = " "
            count = 1
            str = bxInput.Text
            tmAnimate.Enabled = True
            tmDuration.Enabled = True
            'The rest of the action is sent to tmAnimate()
        Else
            lblTag.SendToBack()
            lblDelay.Visible = True
            If pnlBase.BackColor.GetBrightness > 0.4 Then       'Countdown label adaptation
                lblDelay.ForeColor = Color.FromArgb(64, 64, 64)
            Else lblDelay.ForeColor = Color.White
            End If

            lblDelay.Text = wait
            wait -= 1
            If wait = -1 Then
                lblDelay.Visible = False
                lblTag.BringToFront()
                tmDelay.Enabled = False
                bxDisplay.Text = " "
                count = 1
                str = bxInput.Text
                tmAnimate.Enabled = True
                tmDuration.Enabled = True
                'The rest of the action is sent to tmAnimate()
            End If
        End If
    End Sub

    Private Sub tmDone_Tick(sender As Object, e As EventArgs) Handles tmDone.Tick
        If wait = 1 Then        'End operation
            tmDone.Enabled = False
            lblTag.SendToBack()
            PushOrPull()
        Else wait -= 1
        End If
    End Sub

    Dim sec As Integer = 0
    Private Sub tmDuration_Tick(sender As Object, e As EventArgs) Handles tmDuration.Tick
        sec += 1
        lblSec.Text = sec & " seconds"
    End Sub
#End Region

    Private Sub pnlTitleBar_MouseEnter(sender As Object, e As EventArgs) Handles pnlTitleBar.MouseEnter
        ShowMe()
    End Sub

    Private Sub pnlTitleBar_MouseLeave(sender As Object, e As EventArgs) Handles pnlTitleBar.MouseLeave
        tmHide.Enabled = True
    End Sub

#Region "Color Buttons"
    Private Sub btnIndigo_Click(sender As Object, e As EventArgs) Handles btnIndigo.Click
        Dim btn As Button = CType(sender, Button)
        Dim btnN As String = btn.Name
        Me.BackColor = changeColor(btn)

        pnlTitleBar.BackColor = Color.Indigo
        pnlBase.BackColor = Color.Indigo
        btnSideBar.BackColor = Color.Indigo
        btnSideBarR.BackColor = Color.Indigo
        'bxDisplay.BackColor = pnlTitleBar.BackColor
        dBox()
    End Sub

    Private Sub btnSienna_Click(sender As Object, e As EventArgs) Handles btnSienna.Click
        Dim btn As Button = CType(sender, Button)
        Dim btnN As String = btn.Name
        Me.BackColor = changeColor(btn)

        pnlTitleBar.BackColor = Color.Sienna
        pnlBase.BackColor = Color.Sienna
        btnSideBar.BackColor = Color.Sienna
        btnSideBarR.BackColor = Color.Sienna
        'bxDisplay.BackColor = pnlTitleBar.BackColor
        dBox()
    End Sub

    Private Sub btnSkyBlue_Click(sender As Object, e As EventArgs) Handles btnSkyBlue.Click
        Dim btn As Button = CType(sender, Button)
        Dim btnN As String = btn.Name
        Me.BackColor = changeColor(btn)

        pnlTitleBar.BackColor = Color.SkyBlue
        pnlBase.BackColor = Color.SkyBlue
        btnSideBar.BackColor = Color.SkyBlue
        btnSideBarR.BackColor = Color.SkyBlue
        'bxDisplay.BackColor = pnlTitleBar.BackColor
        dBox()
    End Sub

    Private Sub btnLightGreen_Click(sender As Object, e As EventArgs) Handles btnLightGreen.Click
        Dim btn As Button = CType(sender, Button)
        Dim btnN As String = btn.Name
        Me.BackColor = changeColor(btn)

        pnlTitleBar.BackColor = Color.LightGreen
        pnlBase.BackColor = Color.LightGreen
        btnSideBar.BackColor = Color.LightGreen
        btnSideBarR.BackColor = Color.LightGreen
        'bxDisplay.BackColor = pnlTitleBar.BackColor
        dBox()
    End Sub

    Private Sub btnGold_Click(sender As Object, e As EventArgs) Handles btnGold.Click
        Dim btn As Button = CType(sender, Button)
        Dim btnN As String = btn.Name
        Me.BackColor = changeColor(btn)

        pnlTitleBar.BackColor = Color.Gold
        pnlBase.BackColor = Color.Gold
        btnSideBar.BackColor = Color.Gold
        btnSideBarR.BackColor = Color.Gold
        'bxDisplay.BackColor = pnlTitleBar.BackColor
        dBox()
    End Sub

    Private Sub btnMoreColors_Click(sender As Object, e As EventArgs) Handles btnMoreColors.Click
        ColorDialog1.ShowDialog()
        Dim colorR As Integer = ColorDialog1.Color.R
        Dim colorB As Integer = ColorDialog1.Color.B
        Dim colorG As Integer = ColorDialog1.Color.G

        Me.BackColor = findBack(colorR, colorG, colorB)
        pnlTitleBar.BackColor = ColorDialog1.Color
        pnlBase.BackColor = ColorDialog1.Color
        btnSideBar.BackColor = ColorDialog1.Color
        btnSideBarR.BackColor = ColorDialog1.Color
        'bxDisplay.BackColor = pnlTitleBar.BackColor
        dBox()
    End Sub
#End Region

    Public str As String
    Public count As Integer
    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        If bxInput.Text = "" Then
            MsgBox("Please enter some text before proceeding.")
        ElseIf bxInput.Text.Length < 3 Then
            MsgBox("Not enough text. Please enter some more.")
        Else
            bxDisplay.SelectionAlignment = HorizontalAlignment.Center
            lblSec.Text = "0 seconds"
            wait = numDelay.Value
            bxDisplay.Text = ""
            PushOrPull()
            lblDelay.Text = wait
            tmDelay.Enabled = True
            'Next action is sent to tmDelay
        End If
    End Sub

    Private Sub btnStop_Click(sender As Object, e As EventArgs) Handles btnStop.Click
        tmAnimate.Enabled = False
        tmDuration.Enabled = False
    End Sub

    Private Sub numAniInt_ValueChanged(sender As Object, e As EventArgs) Handles numAniInt.ValueChanged
        tmAnimate.Interval = numAniInt.Value
        If bxInput.TextLength > 0 Then
            lblDuration.Text = durationCalc(s:=numAniInt.Value, t:=bxInput.TextLength) & " seconds"
        End If
    End Sub

    Private Sub bxInput_KeyPress(sender As Object, e As KeyPressEventArgs) Handles bxInput.KeyPress
        If bxInput.TextLength > 0 Then
            lblDuration.Text = durationCalc(s:=numAniInt.Value, t:=bxInput.TextLength) & " seconds"
        End If
    End Sub

    Private Sub numDelay_ValueChanged(sender As Object, e As EventArgs) Handles numDelay.ValueChanged
        wait = numDelay.Value
    End Sub

    Private Sub numFontSize_ValueChanged(sender As Object, e As EventArgs) Handles numFontSize.ValueChanged
        bxInput.Font = New System.Drawing.Font(cbxFont.Text, numFontSize.Value, bxInput.SelectionFont.Style)
        bxDisplay.Font = New System.Drawing.Font(cbxFont.Text, numFontSize.Value, bxInput.SelectionFont.Style)
    End Sub

    Private Sub cbxFont_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxFont.SelectedIndexChanged
        bxInput.Font = New System.Drawing.Font(cbxFont.Text, numFontSize.Value, bxInput.SelectionFont.Style)
        bxDisplay.Font = New System.Drawing.Font(cbxFont.Text, numFontSize.Value, bxInput.SelectionFont.Style)
    End Sub

#Region "Music Player"
    Private Sub btnSideBarR_Click(sender As Object, e As EventArgs) Handles btnSideBarR.Click
        PushOrPullM()
    End Sub

    Private Sub pnlDragBase_DragOver(sender As Object, e As DragEventArgs) Handles pnlDragBase.DragOver
        lblDragTop.ForeColor = pnlTitleBar.BackColor
        lblDragBottom.ForeColor = pnlTitleBar.BackColor
    End Sub

    Private Sub pnlDragBase_DragLeave(sender As Object, e As EventArgs) Handles pnlDragBase.DragLeave
        lblDragTop.Visible = False
        lblDragBottom.Visible = False
    End Sub

    Private Sub btnLoadSong_Click(sender As Object, e As EventArgs) Handles btnLoadSong.Click
        If btnLoadSong.Text = "Load" Then
            OpenFileDialog1.ShowDialog()
        End If
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

    End Sub

    Private Sub pnlDragBase_DragDrop(sender As Object, e As DragEventArgs) Handles pnlDragBase.DragDrop
        btnLoadSong.Text = "Checking..."
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        If files.Count > 1 Then
            MsgBox("You can only load a single file")
            btnLoadSong.Text = "Load"
        Else
            If Not (files(0).ToString.EndsWith(".mp3")) Then
                btnLoadSong.Text = "Load"
                MsgBox("The file, " & files(0) & ", is not a music file. Please load a proper music file.")
            Else
                btnLoadSong.Text = "Loading..."
            End If
        End If
    End Sub

    Private Sub pnlDragBase_DragEnter(sender As Object, e As DragEventArgs) Handles pnlDragBase.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
        lblDragTop.Visible = True
        lblDragBottom.Visible = True
    End Sub

#End Region

    Private Sub btnMinimise_MouseHover(sender As Object, e As EventArgs) Handles btnMinimise.MouseHover
        btnMinimise.BackColor = BackColor
    End Sub

    Private Sub btnMinimise_MouseLeave(sender As Object, e As EventArgs) Handles btnMinimise.MouseLeave
        btnMinimise.BackColor = Color.Transparent
    End Sub

End Class