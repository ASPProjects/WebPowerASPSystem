<!--#include file="../part/header.asp"-->
<div id="ct" class="wp cl">
<!--#include file="../part/sider.asp"-->
<div id="mn">
<div id="ProductsShow">
<span class="title">产品展示</span>
<div class="splitline"></div>
<%
If ID <> "" Or Not IsNumeric(ID) Then
    %><!--#include file="../part/cpzs.asp"--><%
Else
    %><!--#include file="../part/cpcaifeng.asp"--><%
End If
%>
</div>
</div></div>
<!--#include file="../part/footer.asp"-->