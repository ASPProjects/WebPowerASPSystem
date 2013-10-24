<!--#include file="./____Core.asp" -->
<%
Dim ID
ID = Request.QueryString("ID")
If ID<>"" Or Not IsNumeric(ID) Then
    Set rs = Server.CreateObject("Adodb.RecordSet")
    sql = "select * from Qianbo_Products where ViewFlag and ID="&ID
    rs.Open sql, conn, 1, 3
    If rs("SeoKeywords") <> "" Then
        SeoKeywords = rs("SeoKeywords")
    Else
        SeoKeywords = rs("ProductName")
    End If
    If rs("SeoDescription") <> "" Then
        SeoDescription = rs("SeoDescription")
    Else
        SeoDescription = rs("ProductName")
    End If
    SeoTitle = rs("ProductName")
    rs.Close
    Set rs = Nothing
ElseIf request.QueryString("SortID") = "" Then
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
<!--#include file="template/SweeLife/page/product.asp"-->