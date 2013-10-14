<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<br />
<%
If InStr(CompanyAdminPurview, "|35,") = 0 Then
    response.Write "<center>您没有管理该模块的权限！</center>"
    response.End
End If
%>
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
<form action="DelContent.Asp?Result=SiteLink" method="post" name="formDel">
  <tr>
    <th width="8">ID</th>
    <th align="left">链接文字</th>
    <th align="left">链接地址</th>
    <th width="60" align="left">优先级别</th>
    <th width="60" align="left">替换次数</th>
    <th width="60" align="left">打开方式</th>
    <th width="30">状态</th>
    <th align="left" width="60">管理</th>
    <th width="28">操作</th>
  </tr>
  <% SetLinkList() %>
  </form>
</table>
<%
Function SetLinkList()
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
    datafrom = "Qianbo_Link"
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
Response.Write "<td nowrap class=""leftrow"">"&rs("ID")&"</td>" & vbCrLf
Response.Write "<td nowrap class=""leftrow""><a href=""LinkEdit.Asp?id="&rs("ID")&"&Result=Modify"">"&rs("Text")&"</a></td>" & vbCrLf
Response.Write "<td nowrap class=""leftrow"">"&rs("Link")&"</td>" & vbCrLf
Response.Write "<td nowrap class=""leftrow"">"&rs("Order")&"</td>" & vbCrLf
Response.Write "<td nowrap class=""leftrow"">" & vbCrLf
If rs("Replace") = 0 Then
    Response.Write "不限"
Else
    Response.Write rs("Replace")
End If
Response.Write "</td>" & vbCrLf
Response.Write "<td nowrap class=""leftrow"">" & vbCrLf
Select Case CStr(LCase(rs("Target")))
    Case "0"
        Response.Write "<font color=""blue"">原窗口</font>"
    Case "1"
        Response.Write "新窗口"
    Case Else
        Response.Write rs("Target")
End Select
Response.Write "</td>" & vbCrLf
Response.Write "<td nowrap class=""leftrow"">" & vbCrLf
Select Case CStr(LCase(rs("State")))
    Case "0"
        Response.Write "<font color=""red"">禁用</font>"
    Case "1"
        Response.Write "启用"
    Case Else
        Response.Write rs("State")
End Select
Response.Write "</td>" & vbCrLf
Response.Write "<td nowrap class=""leftrow""><a href=""LinkEdit.Asp?Result=Add"">添加</a> <a href=""LinkEdit.Asp?id="&rs("ID")&"&Result=Modify"">修改</a></td>" & vbCrLf
Response.Write "<td nowrap class=""centerrow""><input name='selectID' type='checkbox' value='"&rs("ID")&"'></td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
rs.movenext
Wend
Response.Write "<tr>" & vbCrLf
Response.Write "<td colspan='9' nowrap class=""forumRow"" align=""right""><input onClick=""CheckAll(this.form)"" name=""buttonAllSelect"" type=""button"" id=""submitAllSearch"" value=""全选""> <input onClick=""CheckOthers(this.form)"" name=""buttonOtherSelect"" type=""button"" id=""submitOtherSelect"" value=""反选""> <input type=""submit"" name=""batch"" value=""批量生效"" onClick=""return test();""> <input type=""submit"" name=""batch"" value=""批量失效"" onClick=""return test();""> <input name='batch' type='submit' value='删除所选' onClick=""return test();""></td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Else
    response.Write "<tr><td colspan='9' nowrap class=""centerrow"">暂无相关数据</td></tr>"
End If
Response.Write "<tr>" & vbCrLf
Response.Write "<td colspan='9' nowrap class=""leftrow"">" & vbCrLf
Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
Response.Write "<tr>" & vbCrLf
Response.Write "<td class=""leftrow"">共<font color='red'> "&idcount&" </font>条记录 页次：<font color='red'>"&page&"</font></strong>/"&pagec&"&nbsp;每页：<font color='red'>"&pages&"</font>条</td>" & vbCrLf
Response.Write "<td class=""forumRow"" align=""right"">" & vbCrLf
pagenmin = page - pagenc
pagenmax = page+pagenc
If(pagenmin<1) Then pagenmin = 1
If(page>1) Then response.Write ("<a href='"& myself &"Page=1'><font style='font-size: 14px; font-family: webdings'>9</font></a> ")
If(pagenmin>1) Then response.Write ("<a href='"& myself &"Page="& page - (pagenc * 2 + 1) &"'><font style='font-size: 14px; font-family: webdings'>7</font></a> ")
If(pagenmax>pagec) Then pagenmax = pagec
For i = pagenmin To pagenmax
    If(i = page) Then
    response.Write (" <font color='red'>"& i &"</font> ")
Else
    response.Write ("[<a href="& myself &"Page="& i &">"& i &"</a>]")
End If
Next
If(pagenmax<pagec) Then response.Write (" <a href='"& myself &"Page="& page+(pagenc * 2 + 1) &"'><font style='font-size: 14px; font-family: webdings'>8</font></a> ")
If(page<pagec) Then response.Write ("<a href='"& myself &"Page="& pagec &"'><font style='font-size: 14px; font-family: webdings'>:</font></a> ")
Response.Write "跳转到：第 <input name='SkipPage' style=""width: 30px"" onKeyDown='if(event.keyCode==13)event.returnValue=false' onchange=""if(/\D/.test(this.value)){alert('请在跳转栏输入数字！');this.value='"&Page&"';}"" type='text' value='"&Page&"'> 页" & vbCrLf
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