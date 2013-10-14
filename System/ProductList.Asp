<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
If InStr(CompanyAdminPurview, "|12,") = 0 Then
    response.Write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
    response.End
End If
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<%
Dim Result, Keyword, SortID, SortPath
Result = request.QueryString("Result")
Keyword = request.QueryString("Keyword")
SortID = request.QueryString("SortID")
SortPath = request.QueryString("SortPath")

Function PlaceFlag()
    If Result = "Search" Then
        If Keyword<>"" Then
            Response.Write "产品：列表 -> 检索 -> 关键字：<font color='red'>"&Keyword&"</font>"
        Else
            Response.Write "产品：列表 -> 检索 -> 关键字为空(显示全部产品)"
        End If
    Else
        If SortPath<>"" Then
            Response.Write "产品：列表 -> <a href='ProductList.Asp'>全部</a>"
            TextPath(SortID)
        Else
            Response.Write "产品：列表 -> 全部"
        End If
    End If
End Function
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
<form name="formSearch" method="post" action="Search.Asp?Result=Products">
  <tr>
    <th height="22" sytle="line-height:150%">【产品检索及分类查看】</th>
  </tr>
  <tr>
    <td class="forumRow">关键字：<input name="Keyword" type="text" value="<%=Keyword%>" size="20"> <input name="submitSearch" type="submit" value="搜索产品"></td>
  </tr>
  <tr>
    <td class="forumRow"><%PlaceFlag()%></td>
  </tr>
  </form>
</table>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
<form action="DelContent.Asp?Result=Products" method="post" name="formDel">
  <tr>
    <th>ID</th>
    <th>生效</th>
    <th align="left">产品编号</th>
    <th width="200" align="left">产品标题</th>
    <th>状态</th>
    <th>市场价</th>
    <th>批发价</th>
    <th>型号</th>
    <th>库存</th>
    <th>操作</th>
    <th>选择</th>
  </tr>
  <% ProductsList() %>
  </form>
</table>
<%
Function ProductsList()
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
    datafrom = "Qianbo_Products"
    Dim datawhere
    If Result = "Search" Then
        datawhere = "where ProductName like '%" & Keyword &_
        "%' "
    Else
        If SortPath<>"" Then
            datawhere = "where Instr(SortPath,'"&SortPath&"')>0 "
        Else
            datawhere = ""
        End If
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
If rs("ViewFlag") Then
    Response.Write "<td nowrap align='center' class=""forumRow"" width=""50""><a href=""Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Operation=down""><font color='blue'>生效</font></a></td>" & vbCrLf
Else
    Response.Write "<td nowrap align='center' class=""forumRow"" width=""50""><a href=""Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Operation=up""><font color='red'>未生效</font></a></td>" & vbCrLf
End If
Response.Write "<td nowrap class=""forumRow""><input type=""text"" name=""ProductNo"" size=""12"" value="""&rs("ProductNo")&"""></td>" & vbCrLf
Response.Write "<td nowrap title='"&rs("ProductName")&"' class=""forumRow""><input type=""text"" name=""ProductName"" size=""35"" value="""&rs("ProductName")&"""></td>" & vbCrLf
Response.Write "<td nowrap align='center' class=""forumRow"">" & vbCrLf
If rs("NewFlag") Then Response.Write "<font color='red'>新品</font> "
If rs("CommendFlag") Then Response.Write "<font color='green'>推荐</font>"
If rs("NewFlag") = False And rs("CommendFlag") = False Then Response.Write "普通"
Response.Write "</td>"
Response.Write "<td nowrap align='center' class=""forumRow""><input type=""text"" name=""N_Price"" size=""5"" value="""&rs("N_Price")&""" onkeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false""></td>" & vbCrLf
Response.Write "<td nowrap align='center' class=""forumRow""><input type=""text"" name=""P_Price"" size=""5"" value="""&rs("P_Price")&""" onkeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false""></td>" & vbCrLf
Response.Write "<td nowrap align='center' class=""forumRow""><input type=""text"" name=""ProductModel"" size=""5"" value="""&rs("ProductModel")&"""></td>" & vbCrLf
Response.Write "<td nowrap align='center' class=""forumRow""><input type=""text"" name=""Stock"" size=""5"" value="""&rs("Stock")&""" onkeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false""></td>" & vbCrLf
Response.Write "<td align=""center""nowrap class=""forumRow""><a href='ProductEdit.Asp?Result=Add'>添加</a> <a href='ProductEdit.Asp?Result=Modify&ID="&rs("ID")&"'>修改</a><input name=""pro_id"" type=""hidden"" value="""&rs("ID")&"""></td></td>" & vbCrLf
Response.Write "<td nowrap align='center' class=""forumRow""><input name='selectID' type='checkbox' value='"&rs("ID")&"'></td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
rs.movenext
Wend
Response.Write "<tr>" & vbCrLf
Response.Write "<td colspan='11' nowrap align=""right"" class=""forumRow""><input type=""submit"" name=""batch"" value=""批量修改产品参数"" onClick=""return test();""> <input type=""submit"" name=""batch"" value=""批量生效"" onClick=""return test();""> <input type=""submit"" name=""batch"" value=""批量失效"" onClick=""return test();""> <input onClick=""CheckAll(this.form)"" name=""buttonAllSelect"" type=""button"" id=""submitAllSearch"" value=""全选""> <input onClick=""CheckOthers(this.form)"" name=""buttonOtherSelect"" type=""button"" id=""submitOtherSelect"" value=""反选""> <input name='batch' type='submit' value='删除所选' onClick=""return test();""></td>" & vbCrLf
Response.Write "</tr>" & vbCrLf
Else
    response.Write "<tr><td nowrap align='center' colspan='11' class=""forumRow"">暂无产品信息</td></tr>"
End If
Response.Write "<tr>" & vbCrLf
Response.Write "<td colspan='11' nowrap class=""forumRow"">" & vbCrLf
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

Function TextPath(ID)
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "Select * From Qianbo_ProductSort where ID="&ID
    rs.Open sql, conn, 1, 1
    TextPath = " -> <a href=ProductList.Asp?SortID="&rs("ID")&"&SortPath="&rs("SortPath")&">"&rs("SortName")&"</a>"
    If rs("ParentID")<>0 Then TextPath rs("ParentID")
    response.Write(TextPath)
End Function

Function SortText(ID)
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "Select * From Qianbo_ProductSort where ID="&ID
    rs.Open sql, conn, 1, 1
    SortText = rs("SortName")
    rs.Close
    Set rs = Nothing
End Function
%>