<!--#include file="../part/header.asp"-->
<div id="ct" class="wp cl">
<!--#include file="../part/sider.asp"-->
<div id="mn">
<span class="tit">联系我们</span>


<div id="aboutus">
<div style="width:300px; float:left; line-height:25px">
<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from Qianbo_Others where ViewFlag and ID=1"
rs.Open sql, conn, 1, 3
If rs.EOF Then
Response.Write("<div class=""mm"">暂无相关信息</div>")
Else
If ViewNoRight(rs("GroupID"), rs("Exclusive")) Then
Response.Write(rs("Content"))
else
Response.Write("<div class=""mm"" align=""center"">无查看权限</div>")
End If
rs.update
End If
rs.Close
Set rs = Nothing
%>
</div>
<br />
<div style="width:445px; float:left;  overflow:hidden; margin-top:20px; margin-left:10px">
<iframe
id='mapbarframe'
border='0'
vspace='0'
hspace='0'
marginwidth='0'
marginheight='0'
framespacing='0'
frameborder='0'
scrolling='no'
width='445'
height='400'
src='http://searchbox.mapbar.com/publish/template/template1010/index.jsp?CID=1990ch&tid=tid1010&showSearchDiv=1&cityName=%E4%BD%9B%E5%B1%B1%E5%B8%82&nid=MAPBTNTPXOYAZMFBTPHPX&width=445&height=400&infopoi=2&zoom=10&control=1'></iframe>
</div>
</div>
</div></div>
<!--#include file="../part/footer.asp"-->