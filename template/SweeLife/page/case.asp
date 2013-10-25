<!--#include file="../part/header.asp"-->
<div id="ct" class="wp cl">
<!--#include file="../part/sider.asp"-->
<div id="mn">
<span class="tit">工程案例</span>
<div id="cases">
<%
If request.QueryString("SortID") = "" Then
    %><!--#include file="../part/cpcaifeng2.asp"--><%
ElseIf Not IsNumeric(request.QueryString("SortID")) Then
    SeoTitle = "参数错误"
ElseIf conn.Execute("select * from Qianbo_ProductSort Where ViewFlag and  ID="&request.QueryString("SortID")).EOF Then
    SeoTitle = "参数错误"
Else
    %><!--#include file="../part/caseViewC.asp"--><%
End If
%>

</div>
</div></div>
<!--#include file="../part/footer.asp"-->