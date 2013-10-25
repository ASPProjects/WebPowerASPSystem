<div class="pageMainContent">

<div class="pack">

<%
ID = request.QueryString("ID")
If ID = "" Or Not IsNumeric(ID) Then
response.Write "<center>暂无相关信息</center>"
ElseIf conn.Execute("select * from Qianbo_Products Where ViewFlag and  ID="&ID).EOF Then
response.Write "<center>暂无相关信息</center>"
Else

Set rs = server.CreateObject("adodb.recordset")
sql = "select * from Qianbo_Products where ViewFlag and ID="&ID
rs.Open sql, conn, 1, 3
OtherPic = rs("OtherPic")
If Not(IsNull(OtherPic)) Then
OtherPic = Split(OtherPic, "*")
End If
End If
If ViewNoRight(rs("GroupID"), rs("Exclusive")) Then
%>
<div>


<table width="700" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td colspan="2" valign="top"  width="282" style="text-align: center;padding-top: 20px;">

<a title="<%=rs("ProductName")%>"  class="cloud-zoom" rel="adjustX: 10, adjustY:-4"><img src="<%=HtmlSmallPic(rs("GroupID"),rs("BigPic"),rs("Exclusive"))%>" width="500" title="<%=rs("ProductName")%>" id="datu" style="border:1px #cccccc solid; padding:3px" border="0"></a><br />
<br />
<!-- <a href="ProductBuy.Asp?ProductNo=<%=rs("ProductNo")%>" title="立即订购：<%=rs("ProductName")%>"><img src="Images/buy.gif" border="0" /></a>-->

</td>

</tr>
<tr>
<td colspan="2">

<%

set rst6=conn.execute("select  top 1 *  from Qianbo_Products where ID >"& Request.QueryString("ID") &" and ViewFlag   ORDER BY ID")

if not rst6.EOF then

Response.Write("<div style=""font-weight:900""></div>下一个产品：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<a style="" color:#460002"" href='Product.Asp?ID="&rst6.Fields("ID")&"&SortID=2'>"&rst6.Fields("Productname")&"</a><br>")

end if

set rst7=conn.execute("select  top 1 *  from Qianbo_Products where id <"& Request.QueryString("ID") &" and  ViewFlag   ORDER BY id  DESC")

if not rst7.EOF then


Response.Write(" 上一个产品：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<a style="" color:#460002"" href='Product.Asp?ID="&rst7.Fields("ID")&"&SortID=2'>"&rst7.Fields("Productname")&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")

end if

%>
</td>
</tr>
<!--<tr>
<td colspan="2" ><strong style="background:#CCC; display:block; color:#000; padding-left:15px">产品详细介绍：</strong>

<%=ProcessSitelink(rs("Content"))%>
<br />
<br /></td>
</tr>-->

</table>




</div>
<%Else%>
<div class="mm">无查看权限</div>
<%
End If
rs.update
rs.Close
Set rs = Nothing
%>
</div>

</div>


<script type="text/javascript">



function datu(sa){
document.getElementById('datu').src=sa
ddd="uploadfile/20121227/20121227111757505.jpg"
}
var ddd = document.getElementById("datu").href
</script>
















