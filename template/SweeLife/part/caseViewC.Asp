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
<table width="650" border="0" cellspacing="0" cellpadding="0" align="center" style="line-height:25px">
<tr >
<td colspan="2" >
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