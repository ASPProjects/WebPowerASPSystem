<div class="pageMainContent">

<div class="pack">

<%
ID = request.QueryString("ID")
If ID = "" Or Not IsNumeric(ID) Then
    response.Write "<center>暂无相关信息</center>"
ElseIf conn.Execute("select * from Qianbo_Products Where ViewFlag and  ID="&ID).EOF Then
    response.Write "<center>暂无相关信息</center>"
Else
    Set rs = Server.CreateObject("Adodb.RecordSet")
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
<td valign="top"  width="282">

<a title="<%=rs("ProductName")%>"  class="cloud-zoom" rel="adjustX: 10, adjustY:-4"><img src="<%=HtmlSmallPic(rs("GroupID"),rs("BigPic"),rs("Exclusive"))%>" width="300" title="<%=rs("ProductName")%>" id="datu" style="border:1px #cccccc solid; padding:3px" border="0"></a><br />
<br />
<!-- <a href="ProductBuy.Asp?ProductNo=<%=rs("ProductNo")%>" title="立即订购：<%=rs("ProductName")%>"><img src="Images/buy.gif" border="0" /></a>-->

</td>
<td valign="top" width="418" style="line-height:200%; padding-left:20px;">



<div style="width:400px; height:1px; padding:0px; margin:0px; overflow: hidden"></div>

产品名称：<%=rs("ProductName")%><br />
更新时间：<%=FormatDate(rs("Addtime"),13)%><br />
出品单位：美国凯德防水涂料有限公司<br>
型号:<%=rs("ProductModel")%><br>
规格:<%=rs("ProductNo")%>
<br />
<%
attribute1 = rs("attribute1")
attribute1_value = rs("attribute1_value")
If attribute1 <> " and attribute1_value <> " Then
attribute1_1 = Split(attribute1, "§§§")
attribute1_value_1 = Split(attribute1_value, "§§§")
For i = 0 To UBound(attribute1_value_1)
%>
<%=attribute1_1(i)%>：<%=attribute1_value_1(i)%><br />
<%
Next
End If
%>
浏览次数：
<script language="javascript" src="HitCount.Asp?id=<%=rs("ID")%>&LX=Qianbo_Products"></script>
<script language="javascript" src="HitCount.Asp?action=count&LX=Qianbo_Products&id=<%=rs("ID")%>"></script>
<br />
<%If Not(isnull(OtherPic)) Then%>

<strong style="color:red">更多产品图片：</strong><br />
<%
Dim htmlshop
For htmlshop = 1 To UBound(OtherPic)
%>
<a href="#" onmousemove="datu('<%=trim(OtherPic(htmlshop))%>')" title="多图展示：<%=rs("ProductName")%>" rel="lightbox[otherpic]"><img src="<%=trim(OtherPic(htmlshop))%>" width="110" style="border:1px solid #ccc; margin-left:5px; margin-right:5px; margin-top:5px"></a>
<%
Next
End If
%><br />
<%

set rst6=conn.execute("select  top 1 *  from Qianbo_Products where ID >"& Request.QueryString("ID") &" and ViewFlag   ORDER BY ID")

if not rst6.EOF then

Response.Write("<div style=""font-weight:900""></div>下一个产品：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<a style="" color:#460002"" href='ProductView.Asp?ID="&rst6.Fields("ID")&"&SortID=2'>"&rst6.Fields("Productname")&"</a><br>")

end if

set rst7=conn.execute("select  top 1 *  from Qianbo_Products where id <"& Request.QueryString("ID") &" and  ViewFlag   ORDER BY id  DESC")

if not rst7.EOF then


Response.Write(" 上一个产品：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<a style="" color:#460002"" href='ProductView.Asp?ID="&rst7.Fields("ID")&"&SortID=2'>"&rst7.Fields("Productname")&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")

end if

%>


</td>

</tr>
<tr>
<td colspan="2" ><strong style="background:#CCC; display:block; color:#000; padding-left:15px">产品详细介绍：</strong>

<%=ProcessSitelink(rs("Content"))%>
<br />
<br /></td>
</tr>

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
















