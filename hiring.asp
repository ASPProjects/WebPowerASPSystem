<!--#include file="./____Core.asp" -->
<!--#include file="template/default/common/header_common.asp"-->
<div id="ct" class="wp cl">
<!--#include file="template/default/common/sider.asp"-->
<div id="mn">
<span class="tit">招聘信息</span>


<div id="aboutus">
<%

Set rs = server.CreateObject("adodb.recordset")

    sql = "select * from Qianbo_Others where ViewFlag and ID=2"

rs.Open sql, conn, 1, 3
If rs.EOF Then
%>
            <div class="mm">暂无相关信息</div>
            <%
Else
    If ViewNoRight(rs("GroupID"), rs("Exclusive")) Then
%>
            <%=rs("Content")%>
            <%else%>
            <div class="mm" align="center">无查看权限</div>
            <%
End If
rs.update
End If
rs.Close
Set rs = Nothing
%>
</div>
</div></div>
<!--#include file="template/default/common/footer_common.asp"-->