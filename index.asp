<!--#include file="Include/Const.Asp" -->
<!--#include file="Include/NoSQL.Asp" -->
<!--#include file="Include/ConnSiteData.Asp" -->
<!--#include file="dy.Asp" -->
<!--#include file="template/gufeng/page/index.asp"-->
<%
'set conn=server.CreateObject("adodb.connection")
'connstr="DBQ=" + server.mappath(database) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
'conn.Open connstr
'然后在要用conn的时候直接用<!--#include file="conn.asp"-->就可以用了。何必那么麻烦！
'还有你非要那样的话，你就直接用
'sub DB_Connect()
'set conn=server.CreateObject("adodb.connection")
'connstr="DBQ=" + server.mappath(database) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
'conn.Open connstr
'end sub
%>