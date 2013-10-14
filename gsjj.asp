<%
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from Qianbo_about where ViewFlag and ID=1"
rs.Open sql, conn, 1, 3
If rs.EOF Then
Response.Write("<div class=""mm"">暂无相关信息</div>")
Else If ViewNoRight(rs("GroupID"), rs("Exclusive")) Then
	Response.Write(rs("Content"))
else
	Response.Write("<div class=""mm"" align=""center"">无查看权限</div>")
End If
rs.update
End If
rs.Close
Set rs = Nothing
%>
