<!--#include file="../part/header.asp"-->
<div id="ct" class="wp cl">
<!--#include file="../part/sider.asp"-->
<div id="mn">
<div id="News">
<span class="title">新闻中心</span>
<div class="splitline"></div>
<%
If ID <> "" Or Not IsNumeric(ID) Then
    %><!--#include file="../part/xinwen.Asp"--><%
Else
    %><!--#include file="../part/xwcaifeng.Asp"--><%
End If
%>
</div>
</div></div>
<!--#include file="../part/footer.asp"-->