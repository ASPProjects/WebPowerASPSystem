﻿<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="Admin_htmlconfig.Asp"-->
<!--#include file="Fckeditor/Fckeditor.Asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<script language="javascript">
<!--
function setfile(){
    var arr = showModalDialog("eWebEditor/customDialog/file.htm", "", "dialogWidth:22em; dialogHeight:9em; status:0;help=no");
    if (arr ==null){
        alert("系统提示：当前没有上传文件！");
    }
    if (arr !=null){
        editForm.FileUrl.value=arr;
    }
}
//-->
</script>
<%
If InStr(CompanyAdminPurview, "|16,") = 0 Then
    response.Write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
    response.End
End If
Dim Result
Result = request.QueryString("Result")
Dim ID, DownName, ViewFlag, SortName, SortID, SortPath
Dim FileSize, FileUrl, CommendFlag, GroupID, GroupIdName, Exclusive, Content, SeoKeywords, SeoDescription
ID = request.QueryString("ID")
Call DownEdit()
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="editForm" method="post" action="DownEdit.Asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">【<%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>下载】</th>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">下载标题：</td>
      <td width="80%" class="forumRowHighlight"><input name="DownName" type="text" id="DownName" style="width: 280" value="<%=DownName%>" maxlength="100">
        是否生效：<input name="ViewFlag" type="checkbox" value="1" <%if ViewFlag then response.write ("checked")%>>
        是否推荐：<input name="CommendFlag" type="checkbox" value="1" <%if CommendFlag then response.write ("checked")%>> <font color="red">*</font> <input type="button" name="btn" value="复制标题" title="复制标题到：MetaDescription、MetaKeywords" onclick="CopyWebTitle(document.editForm.DownName.value);"></td>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">MetaKeywords：</td>
      <td width="80%" class="forumRowHighlight"><input name="SeoKeywords" type="text" id="SeoKeywords" style="width: 500" value="<%=SeoKeywords%>" maxlength="250"></td>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">MetaDescription：</td>
      <td width="80%" class="forumRowHighlight"><input name="SeoDescription" type="text" id="SeoDescription" style="width: 500" value="<%=SeoDescription%>" maxlength="250"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">下载类别：</td>
      <td class="forumRowHighlight"><input name="SortID" type="text" id="SortID" style="width: 18; background-color:#fffff0" value="<%=SortID%>" readonly> <input name="SortPath" type="text" id="SortPath" style="width: 70; background-color:#fffff0" value="<%=SortPath%>" readonly> <input name="SortName" type="text" id="SortName" value="<%=SortName%>" style="width: 180; background-color:#fffff0" readonly> <a href="javaScript:OpenScript('SelectSort.Asp?Result=Download',500,500,'')">选择类别</a> <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">阅读权限：</td>
      <td class="forumRowHighlight"><select name="GroupID">
          <% call SelectGroup() %>
        </select>
        <input name="Exclusive" type="radio" value="&gt;=" <%if Exclusive="" or Exclusive=">=" then response.write ("checked")%>>
        隶属
        <input type="radio" <%if Exclusive="=" then response.write ("checked")%> name="Exclusive" value="=">
        专属（隶属：权限值≥可查看，专属：权限值＝可查看）</td>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">下载地址：</td>
      <td width="80%" class="forumRowHighlight"><input name="FileUrl" type="text" id="FileUrl" style="width: 280" value="<%=FileUrl%>" maxlength="100"> <input type="button" value="上传文件" onClick="setfile();"> <font color="red">*</font></td>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">文件大小：</td>
      <td width="80%" class="forumRowHighlight"><input name="FileSize" type="text" id="FileSize" style="width: 280" value="<%=FileSize%>" maxlength="100"> MB <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">详细说明：</td>
      <td class="forumRowHighlight"><%
Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath = "FCKeditor/"
oFCKeditor.ToolbarSet = "Default"
oFCKeditor.Width = "100%"
oFCKeditor.Height = "350"
oFCKeditor.Value = Content
oFCKeditor.Create "Content"
%></td>
    </tr>
    <tr>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="保存"> <input type="button" value="返回上一页" onclick="history.back(-1)"></td>
    </tr>
  </form>
</table>
<%
Sub DownEdit()
    Dim Action, rsRepeat, rs, sql
    Action = request.QueryString("Action")
    If Action = "SaveEdit" Then
        Set rs = server.CreateObject("adodb.recordset")
        If Len(Trim(request.Form("DownName")))<3 Then
            response.Write ("<script language='javascript'>alert('请填写下载标题！');history.back(-1);</script>")
            response.End
        End If
        If Request.Form("SortID") = "" And Request.Form("SortPath") = "" Then
            response.Write ("<script language='javascript'>alert('请选择所属分类！');history.back(-1);</script>")
            response.End
        End If
        If Request.Form("FileUrl") = "" Then
            response.Write ("<script language='javascript'>alert('请填写下载地址！');history.back(-1);</script>")
            response.End
        End If
        If Request.Form("FileSize") = "" Then
            response.Write ("<script language='javascript'>alert('请填写下载文件大小！');history.back(-1);</script>")
            response.End
        End If
        If Request.Form("Content") = "" Then
            response.Write ("<script language='javascript'>alert('请填写详细说明！');history.back(-1);</script>")
            response.End
        End If
        If Result = "Add" Then
            sql = "select * from Qianbo_Download"
            rs.Open sql, conn, 1, 3
            rs.addnew
            rs("DownName") = Trim(Request.Form("DownName"))
            If Request.Form("ViewFlag") = 1 Then
                rs("ViewFlag") = Request.Form("ViewFlag")
            Else
                rs("ViewFlag") = 0
            End If
            rs("SortID") = Request.Form("SortID")
            rs("SortPath") = Request.Form("SortPath")
            If Request.Form("CommendFlag") = 1 Then
                rs("CommendFlag") = Request.Form("CommendFlag")
            Else
                rs("CommendFlag") = 0
            End If
            GroupIdName = Split(Request.Form("GroupID"), "┎╂┚")
            rs("GroupID") = GroupIdName(0)
            rs("Exclusive") = Trim(Request.Form("Exclusive"))
            rs("FileSize") = Trim(Request.Form("FileSize"))
            rs("FileUrl") = Trim(Request.Form("FileUrl"))
            rs("Content") = RTrim(Request.Form("Content"))
            rs("SeoKeywords") = Trim(Request.Form("SeoKeywords"))
            rs("SeoDescription") = Trim(Request.Form("SeoDescription"))
            rs("AddTime") = Now()
            rs("UpdateTime") = Now()
            rs.update
            rs.Close
            Set rs = Nothing
            Set rs = server.CreateObject("adodb.recordset")
            sql = "select top 1 ID from Qianbo_Download order by ID desc"
            rs.Open sql, conn, 1, 1
            ID = rs("ID")
            rs.Close
            Set rs = Nothing
            If ISHTML = 1 Then
                Call htmll("", "", ""&DownNameDiy&""&Separated&""&ID&"."&HTMLName&"", "DownView.Asp", "ID=", ID, "", "")
            End If
        End If
        If Result = "Modify" Then
            sql = "select * from Qianbo_Download where ID="&ID
            rs.Open sql, conn, 1, 3
            rs("DownName") = Trim(Request.Form("DownName"))
            If Request.Form("ViewFlag") = 1 Then
                rs("ViewFlag") = Request.Form("ViewFlag")
            Else
                rs("ViewFlag") = 0
            End If
            rs("SortID") = Request.Form("SortID")
            rs("SortPath") = Request.Form("SortPath")
            If Request.Form("CommendFlag") = 1 Then
                rs("CommendFlag") = Request.Form("CommendFlag")
            Else
                rs("CommendFlag") = 0
            End If
            GroupIdName = Split(Request.Form("GroupID"), "┎╂┚")
            rs("GroupID") = GroupIdName(0)
            rs("Exclusive") = Trim(Request.Form("Exclusive"))
            rs("FileSize") = Trim(Request.Form("FileSize"))
            rs("FileUrl") = Trim(Request.Form("FileUrl"))
            rs("Content") = RTrim(Request.Form("Content"))
            rs("SeoKeywords") = Trim(Request.Form("SeoKeywords"))
            rs("SeoDescription") = Trim(Request.Form("SeoDescription"))
            rs("UpdateTime") = Now()
            rs.update
            rs.Close
            Set rs = Nothing
            If ISHTML = 1 Then
                Call htmll("", "", ""&DownNameDiy&""&Separated&""&ID&"."&HTMLName&"", "DownView.Asp", "ID=", ID, "", "")
            End If
        End If
        If ISHTML = 1 Then
            response.Write "<script language='javascript'>alert('设置成功，相关静态页面已更新！');location.replace('DownList.Asp');</script>"
        Else
            response.Write "<script language='javascript'>alert('设置成功！');location.replace('DownList.Asp');</script>"
        End If
    Else
        If Result = "Modify" Then
            Set rs = server.CreateObject("adodb.recordset")
            sql = "select * from Qianbo_Download where ID="& ID
            rs.Open sql, conn, 1, 1
            If rs.bof And rs.EOF Then
                response.Write ("<center>数据库记录读取错误！</center>")
                response.End
            End If
            DownName = rs("DownName")
            ViewFlag = rs("ViewFlag")
            SortName = SortText(rs("SortID"))
            SortID = rs("SortID")
            SortPath = rs("SortPath")
            CommendFlag = rs("CommendFlag")
            GroupID = rs("GroupID")
            Exclusive = rs("Exclusive")
            FileSize = rs("FileSize")
            FileUrl = rs("FileUrl")
            Content = rs("Content")
            SeoKeywords = rs("SeoKeywords")
            SeoDescription = rs("SeoDescription")
            rs.Close
            Set rs = Nothing
        End If
    End If
End Sub

Sub SelectGroup()
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select GroupID,GroupName from Qianbo_MemGroup"
    rs.Open sql, conn, 1, 1
    If rs.bof And rs.EOF Then
        response.Write("未设组别")
    End If
    While Not rs.EOF
        response.Write("<option value='"&rs("GroupID")&"┎╂┚"&rs("GroupName")&"'")
        If GroupID = rs("GroupID") Then response.Write ("selected")
        response.Write(">"&rs("GroupName")&"</option>")
        rs.movenext
    Wend
    rs.Close
    Set rs = Nothing
End Sub

Function SortText(ID)
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "Select * From Qianbo_DownSort where ID="&ID
    rs.Open sql, conn, 1, 1
    SortText = rs("SortName")
    rs.Close
    Set rs = Nothing
End Function
%>