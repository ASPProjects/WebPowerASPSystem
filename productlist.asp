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
<!--#include file="template/default/common/header_common.asp"-->
<div id="ct" class="wp cl">
<!--#include file="template/default/common/sider.asp"-->
<div id="mn">
<span class="tit">产品中心</span>


<div id="aboutus">
<!--#include file="cpcaifeng.asp"-->
</div>
</div></div>
<!--#include file="template/default/common/footer_common.asp"-->