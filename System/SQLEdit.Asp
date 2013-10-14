<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<br />
<%
If InStr(CompanyAdminPurview, "|38,") = 0 Then
    response.Write "<center>您没有管理该模块的权限！</center>"
    response.End
End If
Dim i, tablename
tablename = Trim(request.QueryString("tablename"))
If(Len(tablename)<1) Then
response.Write "<script language='JavaScript'>alert('数据表参数错误！');" & "history.back()" & "</script>"
response.End()
End If
Dim Action, rsCheckAdd, rs, sql
Action = request.QueryString("Action")
If Action = "SaveEdit" Then
    fieldname = Trim(request.Form("fieldname"))
    If(Len(fieldname)<1) Then
    response.Write "<script language='JavaScript'>alert('请输入字段名！');" & "history.back()" & "</script>"
    response.End()
End If
fieldtype = Trim(request.Form("fieldtype"))
If(Len(fieldtype)<1) Then
response.Write "<script language='JavaScript'>alert('请选择字段类型！');" & "history.back()" & "</script>"
response.End()
End If
If fieldtype = "varchar" Then
    charlen = CInt(request.Form("varchar_len"))
    addsql = "alter table "&tablename&" add "&fieldname&" "&fieldtype&" ("&charlen&") null"
Else
    addsql = "alter table "&tablename&" add "&fieldname&" "&fieldtype
End If
conn.Execute(addsql)
Response.Write "<script language=javascript>alert('数据表 "&tablename&" 新字段 "&fieldname&" 添加成功！');window.location.href='"&request.servervariables("http_referer")&"';</script>"
End If
Set rs = server.CreateObject("adodb.recordset")
rs.Open "select top 1 * from "&tablename, conn, 3, 1
%>
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr height="22" sytle="line-height:150%">
    <th align="left" width="200" style="padding-left: 8px">字段名称</th>
    <th align="left" style="padding-left: 8px">字段属性</th>
  </tr>
  <%For i=0 To rs.fields.count-1%>
  <tr>
    <td class="forumRow" style="padding-left: 8px"><%=rs(i).name%></td>
    <td class="forumRow" style="padding-left: 8px"><%
If rs.fields(i).Type = "3" Then
    response.Write "长整型"
    If rs.fields(i).Attributes = "16" Then response.Write " 自动编号字段"
ElseIf rs.fields(i).Type = "202" Then
    response.Write "文本："
    response.Write "长度"&rs.fields(i).DefinedSize
ElseIf rs.fields(i).Type = "2" Then
    response.Write "整形"
ElseIf rs.fields(i).Type = "11" Then
    response.Write "是/否"
ElseIf rs.fields(i).Type = "135" Or rs.fields(i).Type = "7" Then
    response.Write "日期/时间"
ElseIf rs.fields(i).Type = "203" Then
    response.Write "备注"
ElseIf rs.fields(i).Type = "6" Then
    response.Write "货币"
ElseIf rs.fields(i).Type = "205" Then
    response.Write "OLE 对象"
Else
    response.Write "未知"&rs.fields(i).Type
End If
%></td>
  </tr>
  <%Next%>
</table>
<%
rs.Close
Set rs = Nothing
%>
<br />
<script language="javascript">
function seleChan(str){
    if(str=="varchar"){
        document.getElementById("charlen").style.display="";
    }
    else{
        document.getElementById("charlen").style.display="none";
    }
}
</script>
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="form1" method="post" action="?Action=SaveEdit&tablename=<%=tablename%>">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">【添加新字段】</th>
    </tr>
    <tr>
      <td width="200" class="forumRow" align="right">字段名称：</td>
      <td class="forumRowHighlight"><input name="fieldname" type="text" style="width: 150px">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td class="forumRow" align="right">字段类型：</td>
      <td class="forumRowHighlight"><select name="fieldtype" style="width: 100px" onChange="seleChan(this.options[this.selectedIndex].value)">
          <option value="int">长整型</option>
          <option value="smallint">整型</option>
          <option value="varchar">文本</option>
          <option value="datetime">日期/时间</option>
          <option value="memo">备注</option>
          <option value="money">货币</option>
          <option value="bit">是/否</option>
        </select>
        <font color="red">*</font></td>
    </tr>
    <tr id="charlen" style="display:none; float:inherit;">
      <td class="forumRow" align="right">字段长度：</td>
      <td class="forumRowHighlight"><input name="varchar_len" type="text" id="varchar_len" style="width: 100px"></td>
    </tr>
    <tr>
      <td class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="保存设置">
        <input type="button" value="返回上一页" onclick="history.back(-1)"></td>
    </tr>
  </form>
</table>
<br />
<center>
  <font color="red">警告：本功能将直接操作您的数据库结构，不建议非专业人员使用。在进行任何操作前请备份您的数据！</font>
</center>
<br />