Imports System.Data

Partial Class public_RecordDetail
    Inherits PageBase

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
            '取得變數資料
            ID.Text = Request("id")
        End If

        '取得資料
        Dim dtRecord As DataTable = Get_DataSet("RecordDetail", "GetData", ID.Text).Tables(0)

        If dtRecord.Rows.Count > 0 Then
            Dim field_name_list() As String = Split(dtRecord.Rows(0).Item("field_names"), "|")
            Dim data_old_list() As String = Split(dtRecord.Rows(0).Item("data_old"), "_|@")
            Dim data_new_list() As String = Split(dtRecord.Rows(0).Item("data_new"), "_|@")

            For i As Integer = 0 To field_name_list.Length - 1
                Dim NewRow As New TableRow
                Dim NameCell As New TableCell
                Dim FieldCell As New TableCell
                Dim OldCell As New TableCell
                Dim NewCell As New TableCell
                '設定框線
                NameCell.BorderWidth = 1
                FieldCell.BorderWidth = 1
                OldCell.BorderWidth = 1
                NewCell.BorderWidth = 1
                '設定寬度
                NameCell.Width = WebControls.Unit.Parse("14%")
                FieldCell.Width = WebControls.Unit.Parse("15%")
                OldCell.Width = WebControls.Unit.Parse("35%")
                NewCell.Width = WebControls.Unit.Parse("35%")
                '設定資料
                NameCell.Text = Get_Name(dtRecord.Rows(0).Item("prg_code"), field_name_list(i))
                FieldCell.Text = field_name_list(i)
                OldCell.Text = data_old_list(i)
                NewCell.Text = data_new_list(i)
                '加入Row
                NewRow.Cells.Add(NameCell)
                NewRow.Cells.Add(FieldCell)
                NewRow.Cells.Add(OldCell)
                NewRow.Cells.Add(NewCell)
                '加入Table
                Table1.Rows.Add(NewRow)
            Next

        End If
    End Sub

    Protected Function Get_Name(ByVal prgcode As String, ByVal FieldCode As String) As String
        Get_Name = ""

        Dim dt_tmp As DataTable = Get_DataSet("RecordDetail", "GetName", prgcode, FieldCode).Tables(0)
        If dt_tmp.Rows.Count > 0 Then Return dt_tmp.Rows(0).Item(0)
    End Function

End Class
