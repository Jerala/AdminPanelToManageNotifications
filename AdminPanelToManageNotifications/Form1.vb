Imports System.Data.SqlClient
Imports System.Net
Imports System.IO

Structure ParamsTable
    Dim id As Integer
    Dim id_user As Integer
    Dim event_type As Integer
    Dim viewed As Integer
    Dim last_date As Date
    Dim strict_date As Integer
End Structure

Structure IdParams
    Dim id As Integer
    Dim headerText As fontParams
    Dim aText As fontParams
    Dim notes As fontParams
    Dim path_to_scrn As String
    Dim reference As String
End Structure

Structure fontParams
    Dim text As String
    Dim fontSize As Integer
    Dim fontName As String
    Dim bold As Integer
    Dim color As String
    Dim textType As Integer
End Structure

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim ip = FillParams()
        Dim popup As New Form
        popup.StartPosition = FormStartPosition.CenterScreen
        popup.ControlBox = False
        popup.FormBorderStyle = FormBorderStyle.Sizable
        popup.Size = New Size(700, 650)
        popup.BackColor = Color.GhostWhite

        Dim ltblp As New TableLayoutPanel
        ltblp.Dock = DockStyle.Fill
        ltblp.ColumnCount = 10
        ltblp.RowCount = 20
        With ltblp
            With .ColumnStyles
                .Add(New ColumnStyle(SizeType.Percent, 10.0!))
                .Add(New ColumnStyle(SizeType.Percent, 10.0!))
                .Add(New ColumnStyle(SizeType.Percent, 10.0!))
                .Add(New ColumnStyle(SizeType.Percent, 10.0!))
                .Add(New ColumnStyle(SizeType.Percent, 10.0!))
                .Add(New ColumnStyle(SizeType.Percent, 10.0!))
                .Add(New ColumnStyle(SizeType.Percent, 10.0!))
                .Add(New ColumnStyle(SizeType.Percent, 10.0!))
                .Add(New ColumnStyle(SizeType.Percent, 10.0!))
                .Add(New ColumnStyle(SizeType.Percent, 10.0!))
            End With
            With .RowStyles
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
                .Add(New RowStyle(SizeType.Percent, 5.0!))
            End With
        End With

        popup.Controls.Add(ltblp)
        Call CreateAcceptButton(popup, ltblp)
        Call ShowNotification(ltblp, ip)

        popup.ShowDialog()

        If popup.DialogResult = DialogResult.OK Then
            popup.Close()
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If CheckedListBox1.CheckedItems.Count = 0 Then
            MessageBox.Show("Выберите пользователей, которым Вы хотите показывать это сообщение.")
            Exit Sub
        End If
        Dim maxId As Integer = getMaxId()
        Dim pt As New ParamsTable
        pt.id = maxId + 1
        If CheckBox4.Checked Then
            pt.event_type = 2
        Else pt.event_type = 1
        End If
        pt.viewed = 0
        pt.last_date = DateTimePicker1.Value.ToShortDateString
        If RadioButton1.Checked Then
            pt.strict_date = 0
        Else pt.strict_date = 1
        End If
        Dim ip = FillParams()

        Dim cn As New SqlConnection(DefineConStrtoSRV(4))
        Dim cmd = cn.CreateCommand()
        cn.Open()
        For Each itemChecked As DataRowView In CheckedListBox1.CheckedItems
            pt.id_user = DirectCast(CheckedListBox1.Items.Item(CheckedListBox1.Items.IndexOf(itemChecked)), DataRowView).Row.ItemArray(0)
            cmd.CommandText = "INSERT INTO [eventTable] (id, id_user, event_type, viewed, last_date, strict_date) VALUES ('" + pt.id.ToString + "', '" + pt.id_user.ToString + "',
'" + pt.event_type.ToString + "','" + pt.viewed.ToString + "','" + pt.last_date.ToString + "','" + pt.strict_date.ToString + "')"
            cmd.ExecuteNonQuery()
        Next
        cmd.CommandText = "INSERT INTO [eventTableExplanations] (id, headerText, aText, notes, path_to_scrn, reference) VALUES ('" + (maxId + 1).ToString + "', '" + ip.headerText.text + "', '" + ip.aText.text + "',
'" + ip.notes.text + "', '" + ip.path_to_scrn + "', '" + ip.reference + "')"
        cmd.ExecuteNonQuery()
        For i As Integer = 0 To 2
            cmd.CommandText = "INSERT INTO [eventTableExplanations2] (id, fontSize, fontName, bold, color, textType) VALUES ('" + (maxId + 1).ToString + "', '" + ip.headerText.fontSize.ToString + "',
'" + ip.headerText.fontName + "', '" + ip.headerText.bold.ToString + "', '" + ip.headerText.color + "', '" + i.ToString + "')"
            cmd.ExecuteNonQuery()
        Next
        cn.Close()
        cmd.Dispose()
        cn.Dispose()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
    End Sub

    Private Function FillParams() As IdParams
        Dim inte As Integer
        Dim ip As New IdParams
        With ip.headerText
            .text = TextBox1.Text
            If Int32.TryParse(TextBox2.Text, inte) Then
                .fontSize = inte
            Else .fontSize = 14
            End If
            .fontName = TextBox3.Text
            .color = TextBox4.Text
            If CheckBox1.Checked Then
                .bold = 1
            Else .bold = 0
            End If
        End With
        With ip.aText
            .text = TextBox8.Text
            If Int32.TryParse(TextBox7.Text, inte) Then
                .fontSize = inte
            Else .fontSize = 14
            End If
            .fontName = TextBox6.Text
            .color = TextBox5.Text
            If CheckBox2.Checked Then
                .bold = 1
            Else .bold = 0
            End If
        End With
        With ip.notes
            .text = TextBox12.Text
            If Int32.TryParse(TextBox11.Text, inte) Then
                .fontSize = inte
            Else .fontSize = 14
            End If
            .fontName = TextBox10.Text
            .color = TextBox9.Text
            If CheckBox3.Checked Then
                .bold = 1
            Else .bold = 0
            End If
        End With
        ip.path_to_scrn = TextBox13.Text
        ip.reference = TextBox14.Text
        Call SetDefaultValues(ip)
        Return ip
    End Function

    Private Function getMaxId() As Integer
        Dim cn As New SqlConnection(DefineConStrtoSRV(4))
        Dim adap As New SqlDataAdapter("Select TOP 1 id from [dbo].[EventTable] order by id desc", cn)
        Dim dt As New DataTable
        adap.Fill(dt)
        adap.Dispose()
        cn.Dispose()
        Return dt(0)(0)
    End Function

    Function DefineConStrtoSRV(NumOfBonds As Integer, Optional sa As Boolean = False) As String
        DefineConStrtoSRV = ""
        Dim NameBD As String = ""
        Select Case NumOfBonds
            Case 4 : NameBD = "CBR_Emm_Test"
        End Select
        If NameBD <> "" Then
            If sa = True Then
            Else
                DefineConStrtoSRV = "Data Source=212.42.46.12;Initial Catalog=CBR_Emm_Test;Persist Security Info=True;User ID=EmmWarrior;Password=1qwerTY" 'User ID=balans;Password=ygcXTB4hAuSp5c" '"Data Source=212.42.46.12;Initial Catalog=" & NameBD & ";Persist Security Info=True;User ID=EmmWarrior;Password=1qwerTY"
            End If
        End If
    End Function

    Private Sub SetDefaultValues(ByRef ip As IdParams)
        Dim cn As New SqlConnection(DefineConStrtoSRV(4))
        Dim adap As New SqlDataAdapter("Select fontSize, fontName, color from eventTableExplanations2 where id=-1", cn)
        Dim dt As New DataTable
        adap.Fill(dt)
        adap.Dispose()
        cn.Dispose()
        If ip.headerText.fontSize = 0 Then ip.headerText.fontSize = dt(0)(0)
        If ip.headerText.fontName = "" Then ip.headerText.fontName = dt(0)(1)
        If ip.headerText.color = "" Then ip.headerText.color = dt(0)(2)
        If ip.aText.fontSize = 0 Then ip.aText.fontSize = dt(0)(0)
        If ip.aText.fontName = "" Then ip.aText.fontName = dt(0)(1)
        If ip.aText.color = "" Then ip.aText.color = dt(0)(2)
        If ip.notes.fontSize = 0 Then ip.notes.fontSize = dt(0)(0)
        If ip.notes.fontName = "" Then ip.notes.fontName = dt(0)(1)
        If ip.notes.color = "" Then ip.notes.color = dt(0)(2)
    End Sub

    Private Sub FillCheckListBox()
        Dim cn As New SqlConnection("Data Source=212.42.46.12;Initial Catalog=BKO;Persist Security Info=True;User ID=EmmWarrior;Password=1qwerTY")
        Dim adap As New SqlDataAdapter("Select id_usr, FIO_small from USERS", cn)
        Dim dt As New DataTable
        adap.Fill(dt)
        adap.Dispose()
        cn.Dispose()
        With CheckedListBox1
            .DataSource = dt
            .DisplayMember = "FIO_small"
            .ValueMember = "id_usr"
        End With
    End Sub

    Private Sub ShowNotification(popup As TableLayoutPanel, params As IdParams)

        Dim curRow As Integer = 1
        ' Header Text
        If params.headerText.text <> "" Then
            Dim hText As New Label
            hText.Name = "hText"
            hText.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            hText.Text = params.headerText.text
            hText.AutoSize = True
            If params.headerText.bold = 1 Then
                hText.Font = New Font(params.headerText.fontName, params.headerText.fontSize, FontStyle.Bold)
            Else
                hText.Font = New Font(params.headerText.fontName, params.headerText.fontSize, FontStyle.Regular)
            End If
            hText.ForeColor = Color.FromName(params.headerText.color)
            'hText.ForeColor = Color.From
            popup.Controls.Add(hText)
            popup.SetCellPosition(hText, New TableLayoutPanelCellPosition(1, curRow))
            popup.SetRowSpan(hText, 2)
            popup.SetColumnSpan(hText, 8)
            curRow += 2
        End If

        ' aText
        If params.aText.text <> "" Then
            Dim aText As New Label
            aText.Name = "aText"
            aText.Anchor = AnchorStyles.Left Or AnchorStyles.Top
            aText.AutoSize = True
            aText.Text = params.aText.text
            If params.aText.bold = 1 Then
                aText.Font = New Font(params.aText.fontName, params.aText.fontSize, FontStyle.Bold)
            Else
                aText.Font = New Font(params.aText.fontName, params.aText.fontSize, FontStyle.Regular)
            End If
            aText.ForeColor = Color.FromName(params.aText.color)
            popup.Controls.Add(aText)
            popup.SetCellPosition(aText, New TableLayoutPanelCellPosition(1, curRow))
            popup.SetRowSpan(aText, 2)
            popup.SetColumnSpan(aText, 8)
            curRow += 2
        End If

        ' image
        Dim pb As New PictureBox
        pb.Name = "image"
        pb.SizeMode = PictureBoxSizeMode.StretchImage
        pb.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        Try
            Dim req As FtpWebRequest = CType(WebRequest.Create(params.path_to_scrn), FtpWebRequest)
            req.Method = WebRequestMethods.Ftp.DownloadFile
            req.Credentials = New NetworkCredential("Guide", "DjdbyANGCegth777")
            Dim response As FtpWebResponse = req.GetResponse()
            Dim responseStream As Stream = response.GetResponseStream()
            Dim img = Image.FromStream(responseStream)
            pb.Image = img
            pb.Dock = DockStyle.Fill
            popup.Controls.Add(pb)
            popup.SetCellPosition(pb, New TableLayoutPanelCellPosition(1, curRow))
            popup.SetRowSpan(pb, 12)
            popup.SetColumnSpan(pb, 8)
            curRow += 12
        Catch ex As Exception
        End Try

        ' notes
        If params.notes.text <> "" Then
            Dim notes As New Label
            notes.Name = "notes"
            notes.Anchor = AnchorStyles.Left Or AnchorStyles.Top
            notes.Text = params.notes.text
            If params.notes.bold = 1 Then
                notes.Font = New Font(params.notes.fontName, params.notes.fontSize, FontStyle.Bold)
            Else
                notes.Font = New Font(params.notes.fontName, params.notes.fontSize, FontStyle.Regular)
            End If
            notes.ForeColor = Color.FromName(params.notes.color)
            popup.Controls.Add(notes)
            popup.SetCellPosition(notes, New TableLayoutPanelCellPosition(1, curRow))
            popup.SetRowSpan(notes, 2)
            popup.SetColumnSpan(notes, 8)
            curRow += 2
        End If

        ' reference
        If params.reference <> "0" AndAlso params.reference <> "" Then
            Dim btn As New Button
            btn.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
            btn.Width = 160
            btn.Name = "reference"
            btn.Text = "Еще новости по ссылке"
            AddHandler btn.Click, Sub(sender, e) OpenBrowser(params.reference)
            popup.Controls.Add(btn)
            popup.SetCellPosition(btn, New TableLayoutPanelCellPosition(1, curRow))
            popup.SetRowSpan(btn, 1)
            popup.SetColumnSpan(btn, 8)
        End If

    End Sub

    Private Sub OpenBrowser(reference As String)
        Dim frm As New Form
        frm.WindowState = FormWindowState.Maximized
        Dim wb As New WebBrowser
        wb.Dock = DockStyle.Fill
        frm.Controls.Add(wb)
        Try
            Dim uri = New UriBuilder(New Uri(reference))
            If reference.Substring(0, 3) = "ftp" Then
                uri.UserName = "Guide"
                uri.Password = "DjdbyANGCegth777"
            End If
            wb.Navigate(Uri.uri)
            frm.Show()
        Catch ex As Exception
            MessageBox.Show("Не удалось открыть ссылку.")
        End Try
    End Sub

    Private Sub CreateAcceptButton(popup As Form, ltblp As TableLayoutPanel)
        Dim button1 As New Button
        button1.DialogResult = DialogResult.OK
        button1.Text = ""
        button1.Size = New Size(40, 40)
        Dim req As HttpWebRequest = WebRequest.Create("https://cdn4.iconfinder.com/data/icons/evil-icons-user-interface/64/close2-32.png")
        Dim response As HttpWebResponse = req.GetResponse()
        Dim responseStream As Stream = response.GetResponseStream()
        Dim img = Image.FromStream(responseStream)
        button1.BackgroundImage = img
        button1.BackgroundImageLayout = ImageLayout.Stretch
        response.Dispose()
        responseStream.Dispose()

        button1.Name = "OK"
        button1.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        button1.FlatStyle = FlatStyle.Flat
        button1.FlatAppearance.BorderSize = 0
        button1.FlatAppearance.MouseDownBackColor = Color.Transparent
        button1.FlatAppearance.MouseOverBackColor = Color.Transparent
        button1.BackColor = Color.Transparent

        popup.AcceptButton = button1
        ltblp.Controls.Add(button1)
        ltblp.SetCellPosition(button1, New TableLayoutPanelCellPosition(10, 0))
        ltblp.SetRowSpan(button1, 2)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, Button4.Click, Button5.Click
        If FontDialog1.ShowDialog() = DialogResult.OK Then
            Select Case DirectCast(sender, Button).Name
                Case "Button3"
                    TextBox2.Text = Convert.ToString(FontDialog1.Font.Size)
                    TextBox3.Text = Convert.ToString(FontDialog1.Font.Name)
                Case "Button4"
                    TextBox7.Text = Convert.ToString(FontDialog1.Font.Size)
                    TextBox6.Text = Convert.ToString(FontDialog1.Font.Name)
                Case "Button5"
                    TextBox11.Text = Convert.ToString(FontDialog1.Font.Size)
                    TextBox10.Text = Convert.ToString(FontDialog1.Font.Name)
            End Select
        End If
        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            Select Case DirectCast(sender, Button).Name
                Case "Button3"
                    TextBox4.Text = Convert.ToString(ColorDialog1.Color.Name)
                    Try
                        TextBox4.BackColor = Color.FromName(TextBox4.Text)
                    Catch ex As Exception
                        TextBox4.BackColor = ColorTranslator.FromHtml("#" + TextBox4.Text)
                    End Try
                Case "Button4"
                    TextBox5.Text = Convert.ToString(ColorDialog1.Color.Name)
                    Try
                        TextBox5.BackColor = Color.FromName(TextBox5.Text)
                    Catch ex As Exception
                        TextBox5.BackColor = ColorTranslator.FromHtml("#" + TextBox5.Text)
                    End Try
                Case "Button5"
                    TextBox9.Text = Convert.ToString(ColorDialog1.Color.Name)
                    Try
                        TextBox9.BackColor = Color.FromName(TextBox9.Text)
                    Catch ex As Exception
                        TextBox9.BackColor = ColorTranslator.FromHtml("#" + TextBox9.Text)
                    End Try
            End Select
        End If
    End Sub

End Class
