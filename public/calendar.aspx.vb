Imports System.Text.RegularExpressions
Partial Class calendar
    Inherits System.Web.UI.Page
	
    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged
        Dim sScript As String
        Dim sTextBoxID As String
        Dim Syy As String
        Dim Smm As String
        Dim Sdd As String

        '取得要輸入日期的 TextBox  
		Dim reg_exp As New Regex("[\W]+")
		
		sTextBoxID = reg_exp.Replace(Me.Request.QueryString("TextBoxId"),"" )
        'sTextBoxID = Me.Request.QueryString("TextBoxId")
		
		        '將日期設給 TextBox，並將視窗關閉   
        Syy = CStr(Year(Calendar1.SelectedDate.Date))
        Smm = Right("00" & CStr(Month(Calendar1.SelectedDate.Date)), 2)
        Sdd = Right("00" & CStr(Day(Calendar1.SelectedDate.Date)), 2)

        sScript = "opener.window.document.getElementById('" & sTextBoxID & "').value='" & Syy & "/" & Smm & "/" & Sdd & "';"
        sScript = sScript & "window.close();"
        Me.ClientScript.RegisterStartupScript(Me.GetType(), "_Calendar", sScript, True)
    End Sub
End Class
