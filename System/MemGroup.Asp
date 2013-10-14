<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
If InStr(CompanyAdminPurview, "|32,") = 0 Then
    response.Write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
    response.End
End If
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<%
Dim Result
Result = request.QueryString("Result")
Dim ID, GroupID, GroupName, GroupLevel, Explain, AddTime, RanNum
ID = request.QueryString("ID")
Randomize Timer
RanNum = Int((8999) * Rnd + 1009)
If ID = "" Then GroupID = Year(Now)&Month(Now)&Day(Now)&Hour(Now)&Minute(Now)&Second(Now)&RanNum
If Result<>"" Then
    Call MemGroupEdit()
End If
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
<form action="DelContent.Asp?Result=MemGroup" method="post" name="formDel">
  <tr>
    <th>ID</th>
    <th align="left">组别号</th>
    <th align="left">组别名称</th>
    <th align="left">权限值</th>
    <th align="left">说明</th>
    <th align="left">创建时间</th>
    <th>操作</th>
    <th>选择</th>
  </tr>
  <% MemGroupList() %>
  </form>
</table>
<%
Sub MemGroupEdit()
    Dim Action, rs, sql
    Action = request.QueryString("Action")
    If Action = "SaveEdit" Then
        Set rs = server.CreateObject("adodb.recordset")
        If Result = "Add" Then
            sql = "select * from Qianbo_MemGroup"
            rs.Open sql, conn, 1, 3
            rs.addnew
            If Len(Trim(Request.Form("GroupName")))<3 Or Len(Trim(Request.Form("GroupName")))>16 Then
                response.Write "<script language='javascript'>alert('请填写会员组别名称(6-16个字符或3-8个汉字)！');history.back(-1);</script>"
                response.End
            End If
            rs("GroupID") = Request.Form("GroupID")
            rs("GroupName") = Trim(Request.Form("GroupName"))
            rs("GroupLevel") = Trim(Request.Form("GroupLevel"))
            rs("Explain") = Trim(Request.Form("Explain"))
            rs("AddTime") = Now()
        End If
        If Result = "Modify" Then
            sql = "select * from Qianbo_MemGroup where ID="&ID
            rs.Open sql, conn, 1, 3
            If Len(Trim(Request.Form("GroupName")))<3 Or Len(Trim(Request.Form("GroupName")))>16 Then
                response.Write "<script language='javascript'>alert('请填写会员组别名称(6-16个字符或3-8个汉字)！');history.back(-1);</script>"
                response.End
            End If
            rs("GroupName") = Trim(Request.Form("GroupName"))
            rs("GroupLevel") = Trim(Request.Form("GroupLevel"))
            rs("Explain") = Trim(Request.Form("Explain"))
            conn.Execute("Update Qianbo_Members set GroupName='"&Trim(Request.Form("GroupName"))&"' where GroupID='"&Trim(Request.Form("GroupID"))&"'")
        End If
        rs.update
        rs.Close
        Set rs = Nothing
        response.Write "<script language='javascript'>alert('设置成功！');location.replace('MemGroup.Asp');</script>"
    Else
        If Result = "Modify" Then
            Set rs = server.CreateObject("adodb.recordset")
            sql = "select * from Qianbo_MemGroup where ID="& ID
            rs.Open sql, conn, 1, 1
            If rs.RecordCount = 0 Then
                response.Write "<script language='javascript'>alert('无此记录！');history.back(-1)</script>"
                response.End
            End If
            ID = rs("ID")
            GroupID = rs("GroupID")
            GroupName = rs("GroupName")
            GroupLevel = rs("GroupLevel")
            Explain = rs("Explain")
            rs.Close
            Set rs = Nothing
        End If
    End If
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="editMemGroup" method="post" action="MemGroup.Asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>"">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">【修改会员组别】</th>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">ID：</td>
      <td width="80%" class="forumRowHighlight"><input name="ID" type="text" id="ID" style="width: 80" value="<%if ID="" then response.write ("自动") else response.write (ID) end if%>" maxlength="6" readonly></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">组别号：</td>
      <td class="forumRowHighlight"><input name="GroupID" type="text" id="GroupID" style="width: 180" value="<%=GroupID%>" maxlength="18" readonly> <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">组别名称：</td>
      <td class="forumRowHighlight"><input name="GroupName" type="text" id="GroupName" style="width: 180" value="<%=GroupName%>"> <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">权限值：</td>
      <td class="forumRowHighlight"><input name="GroupLevel" type="text" id="GroupLevel" style="width: 80" value="<%=GroupLevel%>" onChange="if(/\D/.test(this.value)){alert('请在权限值中输入整数！');}"> <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">备注：</td>
      <td class="forumRowHighlight"><textarea name="Explain" rows="8" id="Explain" style="width: 500"><%=Explain%></textarea></td>
    </tr>
    <tr>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="保存"></td>
    </tr>
  </form>
</table>
<%
End Sub

Function MemGroupList()
    Dim idCount
    Dim pages
    pages = 20
    Dim pagec
    Dim page
    page = CLng(request("Page"))
    Dim pagenc
    pagenc = 2
    Dim pagenmax
    Dim pagenmin
    Dim datafrom
    datafrom = "Qianbo_MemGroup"
    Dim datawhere
    datawhere = ""
    Dim sqlid
    Dim Myself, PATH_INFO, QUERY_STRING
    PATH_INFO = request.servervariables("PATH_INFO")
    QUERY_STRING = request.ServerVariables("QUERY_STRING")'
    If QUERY_STRING = "" Or InStr(PATH_INFO & "?" & QUERY_STRING, "Page=") = 0 Then
        Myself = PATH_INFO & "?"
    Else
        Myself = Left(PATH_INFO & "?" & QUERY_STRING, InStr(PATH_INFO & "?" & QUERY_STRING, "Page=") -1)
    End If
    Dim taxis
    taxis = "order by id desc"
    Dim i
    Dim rs, sql
    sql = "select count(ID) as idCount from ["& datafrom &"]" & datawhere
    Set rs = server.CreateObject("adodb.recordset")
    rs.Open sql, conn, 0, 1
    idCount = rs("idCount")
    If(idcount>0) Then
    If(idcount Mod pages = 0)Then
    pagec = Int(idcount / pages)
Else
    pagec = Int(idcount / pages) + 1
End If
sql = "select id from ["& datafrom &"] " & datawhere & taxis
Set rs = server.CreateObject("adodb.recordset")
rs.Open sql, conn, 1, 1
rs.pagesize = pages
If page < 1 Then page = 1
If page > pagec Then page = pagec
If pagec > 0 Then rs.absolutepage = page
For i = 1 To rs.pagesize
    If rs.EOF Then Exit For
    If(i = 1)Then
    sqlid = rs("id")
Else
    sqlid = sqlid &","&rs("id")
End If
rs.movenext
Next
End If
If(idcount>0 And sqlid<>"") Then
sql = "select * from ["& datafrom &"] where id in("& sqlid &") "&taxis
Set rs = server.CreateObject("adodb.recordset")
rs.Open sql, conn, 0, 1
While(Not rs.EOF)
Response.Write "<tr>" & vbCrLf
Response.Write "<td nowrap class=""forumRow"">"&rs("ID")&"</td>" & vbCrLf
Response.Write "<td nowrap class=""forumRow"">"&rs("GroupID")&"</td>" & vbCrLf
Response.Write "<td nowrap class=""forumRow"">"&rs("GroupName")&"</td>" & vbCrLf
Response.Write "<td nowrap class=""forumRow""><font color='blue'>"&rs("GroupLevel")&"</font></td>" & vbCrLf
If Len(rs("Explain"))>24 Then
    Response.Write "<td nowrap title='"&rs("Explain")&"' class=""forumRow"">"&Left(rs("Explain"), 22)&"…</td>" & vbCrLf
Else
    Response.Write "<td nowrap title='"&rs("Explain")&"' class=""forumRow"">"&rs("Explain")&"</td>" & vbCrLf
End If
Response.Write "<td nowrap class=""forumRow"">"&rs("AddTime")&"</td>" & vbCrLf
Response.Write "<td align=""center""nowrap class=""forumRow""><a href='MemGroup.Asp?Result=Modify&ID="&rs("ID")&"'>修改</a></td>" & vbCrLf
If rs("ID") = 1 Or rs("ID") = 2 Then
    Response.Write "<td nowrap align='center' class=""forumRow""><font color=""red"">×</font></td>" & vbCrLf
Else
    Response.Write "<td nowrap align='center' class=""forumRow""><input name='selectID' type='checkbox' value='"&rs("GroupID")&"'></td>" & vbCrLf
End If
Response.Write "</tr>" & vbCrLf
rs.movenext
Wend
Response.Write "<tr>" & vbCrLf
Response.Write "<td colspan='10' nowrap align=""right"" class=""forumRow""><input onClick=""CheckAll(this.form)"" name=""buttonAllSelect"" type=""button"" id=""submitAllSearch"" value=""全选""> <input onClick=""CheckOthers(this.form)"" name=""buttonOtherSelect"" type=""button"" id=""submitOtherSelect"" value=""反选""> <input name='submitDelSelect' type='button' id='submitDelSelect' value='删除所选' onClick='ConfirmDel(""是否确定删除？删除后不能恢复！"");'></td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Else
    response.Write "<tr><td nowrap align='center' colspan='10' class=""forumRow"">暂无会员组别</td></tr>"
End If
Response.Write "<tr>" & vbCrLf
Response.Write "<td colspan='10' nowrap class=""forumRow"">" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
Response.Write "<tr>" & vbCrLf
Response.Write "<td class=""forumRow"">共计：<font color='red'>"&idcount&"</font>条记录 页次：<font color='red'>"&page&"</font></strong>/"&pagec&" 每页：<font color='red'>"&pages&"</font>条</td>" & vbCrLf
Response.Write "<td align='right'>" & vbCrLf
pagenmin = page - pagenc
pagenmax = page+pagenc
If(pagenmin<1) Then pagenmin = 1
If(page>1) Then response.Write ("<a href='"& myself &"Page=1'><font style='font-size: 14px; font-family: Webdings'>9</font></a> ")
If(pagenmin>1) Then response.Write ("<a href='"& myself &"Page="& page - (pagenc * 2 + 1) &"'><font style='font-size: 14px; font-family: Webdings'>7</font></a> ")
If(pagenmax>pagec) Then pagenmax = pagec
For i = pagenmin To pagenmax
    If(i = page) Then
    response.Write (" <font color='red'>"& i &"</font> ")
Else
    response.Write ("[<a href="& myself &"Page="& i &">"& i &"</a>]")
End If
Next
If(pagenmax<pagec) Then response.Write (" <a href='"& myself &"Page="& page+(pagenc * 2 + 1) &"'><font style='font-size: 14px; font-family: Webdings'>8</font></a> ")
If(page<pagec) Then response.Write ("<a href='"& myself &"Page="& pagec &"'><font style='font-size: 14px; font-family: Webdings'>:</font></a> ")
Response.Write "第<input name='SkipPage' onKeyDown='if(event.keyCode==13)event.returnValue=false' onchange=""if(/\D/.test(this.value)){alert('请输入需要跳转到的页数并且必须为整数！');this.value='"&Page&"';}"" style='width: 28px;' type='text' value='"&Page&"'>页" & vbCrLf
Response.Write "<input name='submitSkip' type='button' onClick='GoPage("""&Myself&""")' value='转到'>" & vbCrLf
Response.Write "</td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
rs.Close
Set rs = Nothing
Response.Write "</td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
End Function
%>