Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class WebUserMenu
    Inherits System.Web.UI.UserControl
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session
        If Session("usr_name") Is Nothing Or Trim(Session("usr_name")) = "" Then
            Response.Write("<script>alert('連線逾時，請重新登入')</script>")
            Response.Write("<script>setTimeout(window.parent.location.href=""../../Default.aspx"",1000)</script>")

            '若停止javascript則無法自動跳頁至blank.aspx，因此加入下列強制轉換頁面
            '若需要呈現跳出訊息則需註解掉此行程式 by schneider
            Response.Redirect("~/Default.aspx")
            Exit Sub
        End If

        Dim p1 As New PageBase
        Dim ws As New WebReference.LSIO_WebService

        '取得專案名稱
        'Label1.Text = ConfigurationManager.AppSettings("Prj_name")
        usr_name.Text = "使用者:" & Session("usr_name")
        login_time.Text = "　登入時間:" & Session("login_time")
        menu_style.Text = ws.Get_SysPara("css_sys", "01")

        '設定Menu下拉顏色
        Select Case menu_style.Text
            Case "01"
                color.Text = "#3CA48C"
                color2.Text = "#5CC4AC"
                textcolor.Text = "#FFFFFF"
                ontextcolor.Text = "#000"
            Case "02"
                color.Text = "#2C84DB"
                color2.Text = "#4CA4FB"
                textcolor.Text = "#FFFFFF"
                ontextcolor.Text = "#cc0000"
            Case "03"
                color.Text = "#FF9900"
                color2.Text = "#FFBB22"
                textcolor.Text = "#000"
                ontextcolor.Text = "#cc0000"
            Case Else
                color.Text = "#666666"
                color2.Text = "#999999"
        End Select

        '計算未結案件   add by ann
        Dim web As String
        Dim dtMenu As DataTable = New DataTable
        Dim dtList1 As DataTable
        Dim dtList2 As DataTable

        '取得Menu的DataTable
        dtMenu = p1.Get_DataSet("WebUserMenu", "Lv0").Tables(0)

        '功能項目設置
        Dim sub_con As Integer = 8      '子系統數目:從 1 開始
        Dim prg_con As Integer = 20     '每一子系統最大程式數目
        Dim menu(sub_con, prg_con) As String 'menu內容
        web = "../.." 'ws.Get_SysPara("web", "..\..\")

        '動態建立功能項目
        Dim the_limit As String
        Dim sPrg_code As String
        Dim sPrg_name As String
        Dim sIs_group As String
        Dim sPrg_folder As String
        Dim iMaxSize(sub_con) As Integer

        For isub As Integer = 0 To dtMenu.Rows.Count - 1
            '取得第一層List的DataTable
            dtList1 = p1.Get_DataSet("WebUserMenu", "Lv1", dtMenu.Rows(isub).Item("prg_code") & "").Tables(0)

            For ilop As Integer = 0 To dtList1.Rows.Count - 1
                sPrg_code = Trim(dtList1.Rows(ilop).Item("prg_code") & "")
                sPrg_name = Trim(dtList1.Rows(ilop).Item("prg_name") & "")
                sIs_group = Trim(dtList1.Rows(ilop).Item("is_group") & "")
                sPrg_folder = Trim(dtList1.Rows(ilop).Item("prg_folder") & "")

                If sIs_group = "Y" Then
                    '若為群組
                    dtList2 = p1.Get_DataSet("WebUserMenu", "Lv2", sPrg_code).Tables(0)

                    '取得權限字串
                    Dim sTmpCode As String
                    Dim sTmpName As String
                    Dim sTmpFolder As String
                    the_limit = ""
                    For i As Integer = 0 To dtList2.Rows.Count - 1
                        sTmpCode = Trim(dtList2.Rows(i).Item("prg_code") & "")
                        the_limit &= ws.Get_limit(Session("usr_code"), sTmpCode)
                    Next

                    If InStr(the_limit, "E") <> 0 Then
                        '有執行權限
                        Dim iSize As Integer = System.Text.Encoding.Default.GetBytes(sPrg_name).Length
                        If iSize > iMaxSize(isub) Then iMaxSize(isub) = iSize
                        If iSize > 36 Then iMaxSize(isub) = 36
                        If sPrg_name.Length > 36 Then sPrg_name = Mid(sPrg_name, 1, 36) & "<br />" & Mid(sPrg_name, 37)
                        If sPrg_name.Length > 18 Then sPrg_name = Mid(sPrg_name, 1, 18) & "<br />" & Mid(sPrg_name, 19)

                        menu(isub, ilop) = "<li><a href='#'> " & sPrg_name & "</a><ul>"
                        For i As Integer = 0 To dtList2.Rows.Count - 1
                            sTmpCode = Trim(dtList2.Rows(i).Item("prg_code") & "")
                            sTmpName = Trim(dtList2.Rows(i).Item("prg_name") & "")
                            sTmpFolder = Trim(dtList2.Rows(i).Item("prg_folder") & "")
                            the_limit = ws.Get_limit(Session("usr_code"), sTmpCode)

                            If InStr(the_limit, "E") <> 0 Then
                                Select Case sTmpFolder
                                    Case "report"
                                        '報表
                                        menu(isub, ilop) &= "<li><a href='" & web & "/report/" & sTmpCode & ".aspx'> " & sTmpName & "</a></li>"
                                    Case "mboard"
                                        menu(isub, ilop) &= "<li><a href='" & web & "/mboard/" & sTmpCode & ".aspx'> " & sTmpName & "</a></li>"
                                    Case Else
                                        '其他
                                        menu(isub, ilop) &= "<li><a href='" & web & "/" & sTmpFolder & "/" & sTmpCode & "/Default.aspx'> " & sTmpName & "</a></li>"
                                End Select
                            End If
                        Next
                        menu(isub, ilop) &= "</ul></li>"
                    Else
                        ''沒有執行權限
                        'menu(isub, ilop) = "<li><a href='#'> " & sPrg_name & "</a></li>"
                    End If
                Else
                    '取得權限字串
                    the_limit = ws.Get_limit(Session("usr_code"), sPrg_code)

                    If InStr(the_limit, "E") <> 0 Then
                        '有執行權限
                        Dim iSize As Integer = System.Text.Encoding.Default.GetBytes(sPrg_name).Length
                        If iSize > iMaxSize(isub) Then iMaxSize(isub) = iSize
                        If iSize > 36 Then iMaxSize(isub) = 36
                        If sPrg_name.Length > 36 Then sPrg_name = Mid(sPrg_name, 1, 36) & "<br />" & Mid(sPrg_name, 37)
                        If sPrg_name.Length > 18 Then sPrg_name = Mid(sPrg_name, 1, 18) & "<br />" & Mid(sPrg_name, 19)

                        Select Case sPrg_folder
                            Case "report"
                                '報表
                                menu(isub, ilop) = "<li><a href='" & web & "/report/" & sPrg_code & ".aspx'> " & sPrg_name & "</a></li>"

                            Case "mboard"
                                menu(isub, ilop) = "<li><a href='" & web & "/mboard/" & sPrg_code & ".aspx'> " & sPrg_name & "</a></li>"

                            Case Else
                                '其他
                                menu(isub, ilop) = "<li><a href='" & web & "/" & sPrg_folder & "/" & sPrg_code & "/default.aspx'> " & sPrg_name & "</a></li>"

                        End Select
                    Else
                        ''沒有執行權限
                        menu(isub, ilop) = "<li><a href='#'> " & sPrg_name & "</a></li>"
                    End If
                End If
            Next
        Next

        '串成html
        Dim Html_Str As String
        Html_Str = "<div class='nav-container-outer'>"
        Html_Str &= "<img src='" & web & "/images/menu-" & menu_style.Text & "/nav-bg-l.jpg' alt='' class='float-left' />"
        Html_Str &= "<img src='" & web & "/images/menu-" & menu_style.Text & "/nav-bg-r.jpg' alt='' class='float-right' />"

        Html_Str &= "<ul id='nav-container' class='nav-container'>"
        'Html_Str &= "<li> <a class='item-primary' href='" & web & "/new_blank_tree.aspx'>H.首頁</a></li>"

        For i As Integer = 0 To dtMenu.Rows.Count - 1
            Html_Str &= "<li><span class='divider divider-vert'></span></li><li><a class='item-primary' href='#'>" & dtMenu.Rows(i).Item("prg_name") & "</a>"
            Html_Str &= "<ul style='width:" & iMaxSize(i) * 7.5 + 4 & "px;'>"

            For j As Integer = 0 To UBound(menu, 2)
                If menu(i, j) Is Nothing Then
                Else
                    Html_Str = Html_Str & menu(i, j)
                End If
            Next
            Html_Str &= "</ul>"
            Html_Str &= "</li>"
        Next

        Html_Str &= "<li><span class='divider divider-vert'></span></li><li> <a class='item-primary' href='#'>T.工具</a>"
        Html_Str &= "<ul  style='width:100px;'>"
        Html_Str &= "<li><a href=" & web & "/logout.aspx>登出</a></li>"
        Html_Str &= "<li><a href=" & web & "/management/BDP006A/Default.aspx>變更密碼</a></li>"
        'Html_Str &= "<li><a href=" & web & "/management/BDP007A/Default.aspx>參數設定</a></li>"
        Html_Str &= "</ul>"
        Html_Str &= "</li>"
        Html_Str &= "</ul>"
        Html_Str &= "</div>"

        Literal1.Text = Html_Str
    End Sub
End Class
