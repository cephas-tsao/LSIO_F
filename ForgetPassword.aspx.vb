Imports System.Data
Imports System.Text.RegularExpressions
Partial Class ForgetPassword
    Inherits PageBase

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ID As String = RepSql(txt_id.Text)
        Dim Mail As String = RepSql(txt_mail.Text)

        If String.IsNullOrEmpty(ID) Then
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('請填寫您的登入帳號');</script>"))
            Exit Sub
        End If
		
		if  NOt( Regex.Match(ID, "[^0-9a-zA-z.-{7}+]").Success ) then
		  Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('請填寫您的帳號');</script>"))
            
            Exit Sub
        End If
			
		if  NOt( Regex.Match(Mail, "[\w-.]{10+}").Success ) then
		  Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('請填寫輸入信箱');</script>"))
            Exit Sub
        End If
		
		
		
        If String.IsNullOrEmpty(Mail) Then
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('請填寫您登入帳號所設定的電子信箱');</script>"))
            Exit Sub
        End If
        If Chk_Email(Mail) = False Then
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('電子信箱格式錯誤，請確認');</script>"))
            Exit Sub
        End If

        Dim ws As New WebReference.LSIO_WebService
        Dim dt_User As DataTable = Get_DataSet("ForgetPassword", "Info", txt_id.Text).Tables(0)

        If dt_User.Rows.Count > 0 Then
            Dim sName As String = Trim(dt_User.Rows(0).Item(0) & "")
            Dim sMail As String = Trim(dt_User.Rows(0).Item(1) & "")

            If String.IsNullOrEmpty(sMail) Then
                Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('發件人信箱未設定或空白，請洽單位管理人員設定');</script>"))
                Exit Sub
            End If
            If Mail <> sMail Then
                Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('您輸入的電子信箱與此帳號所設定電子信箱不符，請確認');</script>"))
                Exit Sub
            End If

            Dim sPwd As String = Trim(dt_User.Rows(0).Item(2) & "")
            Dim dtMail As New DataTable
            Dim column1 As New DataColumn
            Dim column2 As New DataColumn
            column1.DataType = System.Type.GetType("System.String")
            column2.DataType = System.Type.GetType("System.String")
            dtMail.Columns.Add(column1)
            dtMail.Columns.Add(column2)
            Dim rowMail As DataRow = dtMail.NewRow
            rowMail(0) = sMail
            rowMail(1) = sName
            dtMail.Rows.Add(rowMail)

            Dim sBody As String = "" & _
                "<B>勞工局勞動檢查處-密碼通知信件</B><BR><BR>" & _
                "您的密碼：" & ws.EncodeString(sPwd, "D") & "<BR><BR>" & _
                "<B><font color=""red"">※為確保您個人資訊安全，此信件請勿供他人觀看</font></B>"


            '發送Mail
            If SendMail("勞工局勞動檢查處-密碼通知信件", dtMail, sBody, "ForgetPassword") Then
                Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('發送成功，請至信箱收取信件');</script>"))
            Else
                Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('發送失敗，請聯絡系統管理員');</script>"))
            End If
        Else
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('此帳號不存在或已被停用，請確認');</script>"))
        End If
    End Sub
End Class
