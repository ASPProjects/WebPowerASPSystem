﻿<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
If InStr(CompanyAdminPurview, "|40,") = 0 Then
    response.Write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
    response.End
End If
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<%
Dim Result, Keyword
Result = request.QueryString("Result")
Keyword = request.QueryString("Keyword")

Function PlaceFlag()
    If Result = "Search" Then
        If Keyword <> "" Then
            Response.Write "客户咨询：列表 -> 检索 -> 关键字：<font color='red'>"&Keyword&"</font>"
        Else
            Response.Write "客户咨询：列表 -> 检索 -> 关键字为空(显示全部客户咨询)"
        End If
    Else
        Response.Write "客户咨询：列表 -> 全部"
    End If
End Function
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
<form name="formSearch" method="post" action="Search.Asp?Result=UserMessage">
  <tr>
    <th height="22" sytle="line-height:150%">【客户咨询检索及分类查看】</th>
  </tr>
  <tr>
    <td class="forumRow">关键字：<input name="Keyword" type="text" value="<%=Keyword%>" size="20"> <input name="submitSearch" type="submit" value="搜索客户咨询"></td>
  </tr>
  <tr>
    <td class="forumRow"><%PlaceFlag()%></td>
  </tr>
  </form>
</table>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
<form action="DelContent.Asp?Result=UserMessage" method="post" name="formDel">
  <tr>
    <th width="5%">ID</th>
    <th width="6%">处理</th>
    <th align="left">客户咨询内容</th>
    <th width="10%" align="left">电话</th>
    <th width="10%" align="left">地址</th>
    <th width="10%" align="left">信箱</th>
    <th width="10%" align="left">发布时间</th>
    <th width="6%">选择</th>
  </tr>
  <% UserMessage() %>
  </form>
</table>
<%
Function UserMessage()
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
    datafrom = "Qianbo_Biz"
    Dim datawhere
    If Result = "Search" Then
        datawhere = "where BizContent like '%" & Keyword &_
        "%' "
    Else
        datawhere = ""
    End If
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
    taxis = "order by ID desc"
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
sql = "select * from ["& datafrom &"] where ID in("& sqlid &") "&taxis
Set rs = server.CreateObject("adodb.recordset")
rs.Open sql, conn, 0, 1
Dim SSRnum
SSRnum = 0
While(Not rs.EOF)
Response.Write "<tr onclick=""showDetail("&SSRnum&")"" style=""cursor: hand"">" & vbCrLf
Response.Write "<td nowrap class=""forumRow"">"&rs("ID")&"</td>" & vbCrLf
If rs("BizOK") = 1 Then
    Response.Write "<td nowrap align='center' class=""forumRow"" width=""50""><font color='blue'>处理</font></td>" & vbCrLf
Else
    Response.Write "<td nowrap align='center' class=""forumRow"" width=""50""><font color='red'>未处理</font></td>" & vbCrLf
End If
If StrLen((rs("BizContent"))) > 58 Then
    Response.Write "<td nowrap title='"&rs("BizContent")&"' class=""forumRow"">"&StrLeft(rs("BizContent"), 55)&"</td>" & vbCrLf
Else
    Response.Write "<td nowrap title='"&rs("BizContent")&"' class=""forumRow"">"&rs("BizContent")&"</td>" & vbCrLf
End If
Response.Write "<td nowrap class=""forumRow"">"&rs("BizPhone")&"</td>" & vbCrLf
Response.Write "<td nowrap class=""forumRow"">"&rs("BizAddr")&"</td>" & vbCrLf
Response.Write "<td nowrap class=""forumRow"">"&rs("BizEMail")&"</td>" & vbCrLf
Response.Write "<td nowrap class=""forumRow"">"&rs("BizDate")&"</td>" & vbCrLf
Response.Write "<td nowrap align='center' class=""forumRow""><input name='selectID' type='checkbox' value='"&rs("ID")&"'></td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Response.Write "<tr style=""display:none;"" id=""detail_"&SSRnum&""">" & vbCrLf
Response.Write "<td colspan=""8"" nowrap align=""left"" bgcolor=""#FFFFF0"">"&rs("BizContent")&"<br />"&rs("BizPhone")&" "&rs("BizEMail")&"<br />"&rs("BizAddr")&"</td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
SSRnum = SSRnum + 1
rs.movenext
Wend
Response.Write "<tr>" & vbCrLf
Response.Write "<td colspan='8' nowrap align=""right"" class=""forumRow""><input type=""submit"" name=""batch"" value=""批量处理"" onClick=""return test();""> <input onClick=""CheckAll(this.form)"" name=""buttonAllSelect"" type=""button"" id=""submitAllSearch"" value=""全选""> <input onClick=""CheckOthers(this.form)"" name=""buttonOtherSelect"" type=""button"" id=""submitOtherSelect"" value=""反选""> <input name='batch' type='submit' value='删除所选' onClick=""return test();""></td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Else
    response.Write "<tr><td nowrap align='center' colspan='8' class=""forumRow"">暂无企业客户咨询</td></tr>"
End If
Response.Write "<tr>" & vbCrLf
Response.Write "<td colspan='8' nowrap class=""forumRow"">" & vbCrLf
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
<script language="javascript">
<!--
function showDetail(n)
{
    var o = document.getElementById("detail_"+n);
    o.style.display = o.style.display?"":"none";
}
//-->
</script>