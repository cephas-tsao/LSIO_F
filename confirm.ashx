<%@ WebHandler Language="VB" Class="confirm" %>

'---------------------------------------------------------------------------- 
'程式功能	圖文驗證模組範例 - 產生圖型 
'備註說明	需先產生 Session["confirm"] 
'			範例圖檔僅有 0 ~ 9 的圖型 
'---------------------------------------------------------------------------- 
Imports System
Imports System.Web
Imports System.IO
Imports System.Web.SessionState		' 要使用 Session 必需加入此命名空間

Public Class confirm : Implements IHttpHandler, IRequiresSessionState
    
	Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
		Dim gdata As New MemoryStream()
		Dim bd_img As New BuildImage()
		Dim mdata As String = ""
        
		mdata = context.Session("confirm").ToString()
        
		' 取得驗證圖型的資料（設定驗證圖型尺寸及驗證字串) 
		gdata = bd_img.GenerateImage(200, 54, mdata)
        
		' 設定輸出格式 
		context.Response.ContentType = "image/png"
        
		' 送出資料內容 
		context.Response.BinaryWrite(gdata.ToArray())
	End Sub
 
	Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
		Get
			Return False
		End Get
	End Property

End Class