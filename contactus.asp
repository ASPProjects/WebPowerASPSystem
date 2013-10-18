<!--#include file="./____Core.asp" -->
<%
If ISHTML = 1 Then
    Response.expires = 0
    Response.expiresabsolute = Now() - 1
    Response.addHeader "pragma", "no-cache"
    Response.addHeader "cache-control", "private"
    Response.CacheControl = "no-cache"
End If
ID = request.QueryString("ID")
If ID <> "" Or IsNumeric(ID) Then
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from Qianbo_Others where ViewFlag and ID= 1"
    rs.Open sql, conn, 1, 3
    If rs("SeoKeywords") <> "" Then
        SeoKeywords = rs("SeoKeywords")
    Else
        SeoKeywords = rs("OthersName")
    End If
    If rs("SeoDescription") <> "" Then
        SeoDescription = rs("SeoDescription")
    Else
        SeoDescription = rs("OthersName")
    End If
    SeoTitle = rs("OthersName")
End If
rs.Close
Set rs = Nothing
%>
<!--#include file="template/jielansi/page/contactus.asp"-->