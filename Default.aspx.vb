'---------------------------------------------------------------------------- 
'程式功能	登入畫面
'---------------------------------------------------------------------------- 
Imports System
Imports System.Data.SqlClient
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.Configuration

Partial Public Class _Default
	Inherits System.Web.UI.Page

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
	
		If Not IsPostBack Then
		
			Session("confirm") = getConfirmCode()
		End If
	End Sub

	Public Function getConfirmCode() As String
		Dim rnd As Random = New Random()
		Dim cnt As Integer = 0
		Dim confirm As String = ""

		' 隨機產生四碼的驗證數字 
		For cnt = 0 To 3
			confirm &= rnd.Next(10).ToString()
		Next

		Return confirm
	End Function

	Protected Sub bn_new_confirm_Click(ByVal sender As Object, ByVal e As EventArgs)
		Session("confirm") = getConfirmCode()

		' 使用 Opera 瀏覽器時，要呼叫此 javascript 更新方式，才會更新驗證圖檔的顯示 
		lt_show.Text = "<script language=javascript>renew_img();</script>"
	End Sub

	Protected Sub bn_reset_Click(ByVal sender As Object, ByVal e As EventArgs)
		
		tb_confirm.Text = ""
		Session("confirm") = getConfirmCode()
	End Sub

	Protected Sub bn_ok_Click(ByVal sender As Object, ByVal e As EventArgs)
		Dim cfc As New Common_Func()
		Dim sfc As New String_Func()
		Dim mg_id As String = "", mg_pass As String = "", confirm As String = "", mErr As String = "", tmpstr As String = ""
		Dim tmparray As String()
		Dim strsplit As String() = New String() {vbTab & vbLf}
		' 分隔分辦用字串 
	
		confirm = tb_confirm.Text.Trim()

		If mg_id = "" Then
			mErr &= "請填寫「帳號」!\n"
		End If

		If mg_pass = "" Then
			mErr &= "請填寫「密碼」!\n"
		End If

		If Session("confirm") Is Nothing Then
			mErr &= "驗證碼無法確認!\n"
		ElseIf confirm <> Session("confirm").ToString() Then
			mErr &= "驗證碼輸入錯誤!\n"
		End If

		If mErr = "" Then
			tmpstr = cfc.Check_ID(mg_id, mg_pass, Request.ServerVariables("REMOTE_ADDR"))

			If sfc.Left(tmpstr, 1) = "*" Then
				mErr = tmpstr.Substring(1)
			Else
				tmparray = tmpstr.Split(strsplit, StringSplitOptions.None)

				Session("mg_sid") = tmparray(0)
				Session("mg_name") = tmparray(1)
				Session("mg_power") = tmparray(2)
			End If
		End If

		If mErr = "" Then
			' 全部驗證都正確 
			' 清除驗證碼的 Session 值 
			Session.Remove("confirm")

			' 重新導向至主畫面 
			'Response.Redirect("Main.html")
			Response.Redirect("9001\9001.aspx")
		
		Else
			' 有錯誤 
			' 重新產生驗證碼 
			bn_reset_Click(sender, e)

			' 利用 javascript 顯示錯誤訊息 
			lt_show.Text = "<script language=javascript>alert(""" & mErr & """);</script>"
	
	
		End If
	End Sub
End Class
