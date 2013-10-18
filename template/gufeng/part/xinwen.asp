<TABLE cellSpacing=0 cellPadding=0 width=710 border=0>
<TBODY>
<TR>
<TD height=35>
</TD></TR>
<TR>
<TD
style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; LINE-HEIGHT: 25px; PADDING-TOP: 10px; color:#b4a29">
<TABLE cellSpacing=0 cellPadding=2 width="100%"
align=center border=0>
<TBODY>
<TR>
<TD align=middle height=28><B><SPAN style="FONT-SIZE: 14px; text-align:center; color:#b4a29"><%=SeoTitle%>
</SPAN></B></TD></TR>

<TR>
<TD align=middle>


<div class="pack">
<div class="ml"></div>
<div class="mr"></div>
<%
ID = request.QueryString("ID")
If ID = "" Or Not IsNumeric(ID) Then
response.Write "<center>暂无相关信息</center>"
ElseIf conn.Execute("select * from Qianbo_News Where ViewFlag and  ID="&ID).EOF Then
response.Write "<center>暂无相关信息</center>"
Else
Dim rs, sql
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from Qianbo_News where ViewFlag and ID="&ID
rs.Open sql, conn, 1, 3
End If
If Request("page") = "" Then
pageNum = 0
Else
pageNum = Request("page")
End If
If ViewNoRight(rs("GroupID"), rs("Exclusive")) Then
%>
<div class="mm">发布时间：<font color="#666"><%=FormatDate(rs("Addtime"),13)%></font> 新闻来源：<font color="#666"><%=rs("Source")%></font> 浏览次数：<font color="#666">
<script language="javascript" src="HitCount.Asp?id=<%=rs("ID")%>&LX=Qianbo_News"></script>
<script language="javascript" src="HitCount.Asp?action=count&LX=Qianbo_News&id=<%=rs("ID")%>"></script>
</font></div>
<div class="mm" style=" text-align:left">
<%
ContentStr = Split(rs("Content"), "{＄html_Paging＄}")
For i = pageNum To pageNum
%>
<%=ProcessSitelink(ContentStr(i))%>
<% Next %>
</div>
<div class="mm">本文共分
<%For p = 0 to ubound(ContentStr)%>
<a href="NewsView.Asp?ID=<%=request("ID")%>&page=<%=p%>"><font color="red"><%=p+1%></font></a>
<%Next%>
页</div>
<%

set rst6=conn.execute("select  top 1 *  from Qianbo_News where ID >"& Request.QueryString("ID") &" and ViewFlag  and SortID="&rs("SortID")&"  ORDER BY ID")

if not rst6.EOF then

Response.Write("<div style=""font-weight:900""></div>上一篇：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<a style="" color:#666"" href='NewsView.Asp?ID="&rst6.Fields("ID")&"&SortID=2'>"&rst6.Fields("NewsName")&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")

end if

set rst7=conn.execute("select  top 1 *  from Qianbo_News where id <"& Request.QueryString("ID") &" and  ViewFlag and SortID="&rs("SortID")&"   ORDER BY id  DESC")

if not rst7.EOF then


Response.Write(" 下一篇：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<a style="" color:#666"" href='NewsView.Asp?ID="&rst7.Fields("ID")&"&SortID=2'>"&rst7.Fields("NewsName")&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")

end if

%>
<%Else %>
<div class="mm">无查看权限</div>
<%
End If
rs.update
rs.Close
Set rs = Nothing
%>

</div>

</TD>
</TR>

<TR>
<TD align=right height=25>&nbsp;</TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>