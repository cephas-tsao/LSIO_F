Imports System.Data
Imports System.Xml

Partial Class UserIdCheck
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "UserIdCheck"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
            Dim SSO_ws As New SSO_Service.ap
            Dim Dt_db As DataTable = Get_DataSet("Login", "IDno").Tables(0)

            For i As Integer = 0 To Dt_db.Rows.Count - 1
                If Trim(Dt_db.Rows(i).Item("is_use") & "") = "Y" Then
                    Dim result As String = ""
                    Dim user_idno As String = ws.EncodeString(Trim(Dt_db.Rows(i).Item("usr_idno") & ""), "D")
                    Dim Xml As XmlNode = SSO_ws.GetUserBasicInfo(user_idno, "LSIOServiceTest")

                    For j As Integer = 0 To Xml.ChildNodes.Count - 1
                        'Response.Write(Xml.ChildNodes(j).Name & ":" & Trim(Xml.ChildNodes(j).InnerText) & "<br>")

                        Select Case Xml.ChildNodes(j).Name
                            Case "result"
                                result = Trim(Xml.ChildNodes(j).InnerText)
                        End Select
                    Next
                    'Response.Write("result:" & result & "<br>")

                    If result = "true" Then
                        If Trim(Dt_db.Rows(i).Item("count") & "") <> "0" Then SaveEditData(s_prgcode, "True", Trim(Dt_db.Rows(i).Item("usr_code") & ""))
                    Else
                        SaveEditData(s_prgcode, "False", Trim(Dt_db.Rows(i).Item("usr_code") & ""))
                    End If
                End If
            Next

            '停用超過5週無權限使用者資料
            SaveEditData(s_prgcode, "StopUse")
        End If
    End Sub
End Class
