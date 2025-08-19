Imports System.Data

Partial Class basic_LSI016A_preview
    Inherits PageBase

    Dim s_prgcode As String = "LSI016A"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.CacheControl = "no-cache"
        Response.AddHeader("Pragma", "no-cache")
        Response.Expires = -1
        Response.Buffer = True

        Dim sNewsCode As String = Request.QueryString("code")
        Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", sNewsCode).Tables(0)
        Dim sNewsTitle As String = Trim(Dt_tmp.Rows(0).Item("news_title") & "")
        Dim sNewsType As String = Trim(Dt_tmp.Rows(0).Item("news_type") & "")
        Dim sNewsContect As String = Trim(Dt_tmp.Rows(0).Item("news_contect") & "")
        Dim sNewsPic1 As String = Trim(Dt_tmp.Rows(0).Item("news_pic1") & "")
        '照片網站位址
        Dim sWebPic As String = Get_SysPara("web") & "/FileUpload/File_LSI016A/"

        Dim sHtml As String = "" & _
            "<html xmlns='http://www.w3.org/1999/xhtml'>" & _
            "<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8' />" & _
            "<style type='text/css'>" & _
            "body{margin-left: 0px;margin-top:15px;margin-right:0px;margin-bottom:0px;}" & _
            ".title1{font-family:微軟正黑體;font-size:x-large;color:navy;font-weight:600;line-height:normal;text-decoration:underline}" & _
            ".title2{font-family:微軟正黑體;font-size:large;color:red;font-weight:600;line-height:normal;text-decoration:underline}" & _
            ".context{font-family:微軟正黑體;font-size:x-large;color:blue;font-weight:600;line-height:normal;text-decoration:underline}" & _
            "</style></head><body>" & _
            "<table cellpadding=0 cellspacing=0 border=0  style='position:relative; margin-left:auto; margin-right:auto; width:990px;' summary='排版表格'>"


        Select Case sNewsType
            Case "01" '電子報明細-最新消息
                sHtml &= "<tr><td><br>"
                '新聞標題
                If sNewsContect <> "" Then
                    sHtml &= "<p class='title2'>" & sNewsTitle & "</p>"
                End If

                If sNewsPic1 <> "" Then
                    '照片
                    sHtml &= "<table cellspacing=0 style= 'width:276px;height:180px;' cellpadding=3 align=left cellspacing=0><tr><td>" & _
                            "<img alt='照片' src='" & sWebPic & sNewsPic1 & "' height=216 width=313></td></tr></table>"
                End If
                '內文
                sHtml &= "<p>" & Replace(sNewsContect, "", "&nbsp;") & "</p></td></tr>"
                '分隔線
                sHtml &= "<tr style='height:1px'><td><hr size=1 noshade style='border:1px dashed #cccccc;'><td></tr>"

            Case "02" '電子報明細-交流園地
                sHtml &= "<tr><td><br>"
                '新聞標題
                If sNewsContect <> "" Then
                    sHtml &= "<p class='title2'><span class='style26'>" & sNewsTitle & "</span></p>"
                End If

                If sNewsPic1 <> "" Then
                    '照片
                    sHtml &= "<table cellspacing=0 style= 'width:155px;height:102px;' cellpadding=3 align=left cellspacing=0><tr><td>" & _
                            "<img alt='照片' src='" & sWebPic & sNewsPic1 & "' height=102 width=153></td></tr></table>"
                End If
                '內文
                sHtml &= "<p>" & Replace(sNewsContect, "", "&nbsp;") & "</p></td></tr>"
                '分隔線
                sHtml &= "<tr style='height:1px'><td><hr size=1 noshade style='border:1px dashed #cccccc;'><td></tr>"

            Case "03" '電子報明細-法規新訊
                sHtml &= "<tr><td><br>"
                '新聞標題
                If sNewsContect <> "" Then
                    sHtml &= "<p class='title2'>" & sNewsTitle & "</p>"
                End If

                If sNewsPic1 <> "" Then
                    '照片
                    sHtml &= "<table cellspacing=0 style= 'width:153px;height:102px;' cellpadding=3 align=left cellspacing=0><tr><td>" & _
                            "<img alt='照片' src='" & sWebPic & sNewsPic1 & "' height=102 width=153></td></tr></table>"
                End If
                '內文
                sHtml &= "<p>" & Replace(sNewsContect, "", "&nbsp;") & "</p></td></tr>"
                '分隔線
                sHtml &= "<tr style='height:1px'><td><hr size=1 noshade style='border:1px dashed #cccccc;'><td></tr>"

            Case "04" '電子報明細-訊息報導
                sHtml &= "<tr><td><br>"
                '新聞標題
                If sNewsContect <> "" Then
                    sHtml &= "<p class='title2'>" & sNewsTitle & "</p>"
                End If

                If sNewsPic1 <> "" Then
                    '照片
                    sHtml &= "<table cellspacing=0 style= 'width:153px;height:102px;' cellpadding=3 align=left cellspacing=0><tr><td>" & _
                            "<img alt='照片' src='" & sWebPic & sNewsPic1 & "' height=102 width=153></td></tr></table>"
                End If
                '內文
                sHtml &= "<p>" & Replace(sNewsContect, "", "&nbsp;") & "</p></td></tr>"
                '分隔線
                sHtml &= "<tr style='height:1px'><td><hr size=1 noshade style='border:1px dashed #cccccc;'><td></tr>"

            Case "05" '電子報明細-充電小站
                sHtml &= "<tr><td>"
                '新聞標題
                'If sNewsContect <> "" Then
                '    sHtml &= "<p class='style22'>" & sNewsTitle & "</p>"
                'End If

                If sNewsPic1 <> "" Then
                    '照片
                    sHtml &= "<table cellspacing=0 style= 'width:102;height:102px;' cellpadding=3 align=left cellspacing=0><tr><td >" & _
                            "<img alt='照片' src='" & sWebPic & sNewsPic1 & "' height=102 width=153></td></tr></table>"
                End If
                '內文
                sHtml &= "<p>" & Replace(sNewsContect, "", "&nbsp;") & "</p></td></tr>"
                '分隔線
                sHtml &= "<tr style='height:1px'><td><hr size=1 noshade style='border:1px dashed #cccccc;'><td></tr>"

            Case "06" '電子報明細-職災省思
                sHtml &= "<tr><td><br>"
                '新聞標題
                If sNewsContect <> "" Then
                    sHtml &= "<p class='title2'>" & sNewsTitle & "</p>"
                End If

                If sNewsPic1 <> "" Then
                    '照片
                    sHtml &= "<table cellspacing=0 style= 'width:153px;height:116px;' cellpadding=3 align=left cellspacing=0><tr><td>" & _
                            "<img alt='照片' src='" & sWebPic & sNewsPic1 & "' height=102 width=153></td></tr></table>"
                End If
                '內文
                sHtml &= "<p>" & Replace(sNewsContect, "", "&nbsp;") & "</p></td></tr>"
                '分隔線
                sHtml &= "<tr style='height:1px'><td><hr size=1 noshade style='border:1px dashed #cccccc;'><td></tr>"
        End Select
        sHtml &= "</body>"
        Response.Write(sHtml)
    End Sub
End Class
