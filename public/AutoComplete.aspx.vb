Imports System.Collections.Generic
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Data
Imports System.IO

Partial Class public_AutoComplete
    Inherits PageBase

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.CacheControl = "no-cache"
        '使用者目前輸入的文字預設以q傳入
        Dim q As String = Request("q")
        If InStr(q, "'") > 0 Or InStr(q, "[") > 0 Then
            Exit Sub
        End If
        If q.Length > 0 Then
            Dim t As DataTable = getStockData()

            If t.Rows.Count > 0 Then
                Dim dv As New DataView(t)
                '利用LIKE做查詢
                dv.RowFilter = "data LIKE '" & q & "%'"
                dv.Sort = "data"
                'dv.Sort = "Key, Symbol"
                Dim lst As New List(Of String)()
                lst.Add("")
                For Each drv As DataRowView In dv
                    Dim r As DataRow = drv.Row
                    '組裝出前端要用的欄位
                    lst.Add(String.Format("{0}", r("data")))
                    If lst.Count >= 100 Then
                        Exit For
                    End If
                Next
                '每筆資料間以換行分隔
                Response.Write(String.Join(vbLf, lst.ToArray()))
            End If
        End If

        Dim cacheKeys As New ArrayList
        Dim cacheEnum As IDictionaryEnumerator = Cache.GetEnumerator()
        While cacheEnum.MoveNext()
            cacheKeys.Add(cacheEnum.Key.ToString())
        End While

        For Each cacheKey As String In cacheKeys
            Cache.Remove(cacheKey)
        Next
    End Sub

    Private Function getStockData() As DataTable
        Dim dt_tmp1 As DataTable = Get_DataSet("AutoComplete", "", Request("c"), Request("s")).Tables(0)
        If dt_tmp1.Rows.Count > 0 Then Return Get_DataSet("AutoComplete", "SqlStr", dt_tmp1.Rows(0).Item(0)).Tables(0)
        Return New DataTable
    End Function

End Class

