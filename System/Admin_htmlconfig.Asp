<!--#include file="CheckAdmin.Asp"-->
<!--#include file="Admin_html_function.Asp"-->
<%
If InStr(CompanyAdminPurview, "|34,") = 0 Then
    response.Write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
    response.End
End If
server.scripttimeout = 999
Call SiteInfo()

Function htmll(mulu, htmlmulu, FileName, filefrom, htmla, htmlb, htmlc, htmld)
    Call SiteInfo()
    If mulu = "" Then mulu = ""&SysRootDir&""
    If htmlmulu = "" Then htmlmulu = ""&SysRootDir&""
    mulu = Replace(mulu, "//", "/")
    FilePath = Server.MapPath(mulu)&"\"&FileName
    Do_Url = "http://"
    Do_Url = Do_Url&Request.ServerVariables("server_name")&htmlmulu&filefrom
    Do_Url = Do_Url&"?"&htmla&htmlb&"&"&htmlc&htmld
    strUrl = Do_Url
    Set objXmlHttp = Server.CreateObject("Microsoft.XMLHTTP")
    objXmlHttp.Open "GET", strUrl, False
    objXmlHttp.send()
    binFileData = objXmlHttp.responseBody
    Set objXmlHttp = Nothing
    Set objAdoStream = Server.CreateObject("Adodb." & "Stream")
    objAdoStream.Type = 1
    objAdoStream.Open()
    objAdoStream.Write(binFileData)
    objAdoStream.SaveToFile FilePath, 2
    objAdoStream.Close()
    Set objAdoStream = Nothing
End Function
%>