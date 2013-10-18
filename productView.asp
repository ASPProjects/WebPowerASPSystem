<!--#include file="Include/Const.Asp" -->
<!--#include file="Include/NoSQL.Asp" -->
<!--#include file="Include/ConnSiteData.Asp" -->
<!--#include file="dy.Asp" -->
<%
Call SiteInfo()
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
End If
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>
<% =SeoTitle %>
-
<% =SiteTitle %>
</title>
<meta name="keywords" content="<% =SeoKeywords %>" />
<meta name="description" content="<% =SeoDescription %>" />
<link href="css/style.css" rel="stylesheet" type="text/css">
</head>

<body>
<!--#include file="top.asp"-->
<div id="content2">
<div id="con-left">
<div id="renzheng">
<div class="rz-title">国际认证</div>
<div class="rz-neirong"><img src="images/energystar.png"></div>
</div>
<div id="chanpin">
<div class="rz-title">服务项目</div>
<div id="fw-neirong">
<ul>
<!--#include file="left.asp"-->
</ul>
</div>
</div>
</div>
<div id="con-right">
<div id="ab-title">产品展示</div>
<div id="ab-neirong"><!--#include file="cpzs.asp"--></div>
</div>
</div>
<!--#include file="end.asp"-->
</body>
</html>
