<!--#include file="./____Core.asp" -->
<%
If request.QueryString("ID") = "" Then
    SeoTitle = "公司简介"
ElseIf Not IsNumeric(request.QueryString("ID")) Then
    SeoTitle = "参数错误"
ElseIf conn.Execute("select * from Qianbo_About Where ViewFlag and  ID="&request.QueryString("ID")).EOF Then
    SeoTitle = "参数错误"
Else
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from Qianbo_About where ViewFlag and ID="&request.QueryString("ID")
    rs.Open sql, conn, 1, 1
    SeoTitle = rs("AboutName")
    rs.Close
    Set rs = Nothing
End If
%>
<!--#include file="template/gufeng/page/aboutus.asp"-->