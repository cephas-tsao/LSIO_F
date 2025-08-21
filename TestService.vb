Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

<WebService(Namespace:="http://tempuri.org/")>
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Public Class TestService
    Inherits System.Web.Services.WebService

    <WebMethod()>
    Public Function Hello() As String
        Return "Hello world"
    End Function
End Class
