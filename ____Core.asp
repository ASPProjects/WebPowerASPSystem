<!--#include file="Include/Const.Asp" -->
<!--#include file="Include/NoSQL.Asp" -->
<!--#include file="Include/ConnSiteData.Asp" -->
<!--#include file="Include/dy.Asp" -->
<%
Dim SysRoot
SysRoot = Server.MapPath("./")
Call SiteInfo()
If ISHTML = 1 Then
    Response.Expires = 0
    Response.ExpiresAbsolute = Now() - 1
    Response.addHeader "Pragma", "no-cache"
    Response.addHeader "Cache-Control", "Private"
    Response.CacheControl = "No-Cache"
End If
%>