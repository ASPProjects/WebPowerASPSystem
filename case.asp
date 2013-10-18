<!--#include file="./____Core.asp" -->
<%
If request.QueryString("SortID") = "" Then
    SeoTitle = "产品展示"
ElseIf Not IsNumeric(request.QueryString("SortID")) Then
    SeoTitle = "参数错误"
ElseIf conn.Execute("select * from Qianbo_ProductSort Where ViewFlag and  ID="&request.QueryString("SortID")).EOF Then
    SeoTitle = "参数错误"
Else
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from Qianbo_ProductSort where ViewFlag and ID="&request.QueryString("SortID")
    rs.Open sql, conn, 1, 1
    SeoTitle = rs("SortName")
    rs.Close
    Set rs = Nothing
End If
%>
<!--#include file="template/jielansi/page/case.asp"-->