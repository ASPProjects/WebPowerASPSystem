<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="Admin_htmlconfig.Asp"-->
<!--#include file="Fckeditor/Fckeditor.Asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<%
If InStr(CompanyAdminPurview, "|21,") = 0 Then
    response.Write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
    response.End
End If
Dim Result
Result = request.QueryString("Result")
Dim ID, OthersName, ViewFlag, SortName, SortID, SortPath
Dim GroupID, GroupIdName, Exclusive, Content, SeoKeywords, SeoDescription
ID = request.QueryString("ID")
Call OthersEdit()
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="editForm" method="post" action="OthersEdit.Asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">【<%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>信息】</th>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">信息标题：</td>
      <td width="80%" class="forumRowHighlight"><input name="OthersName" type="text" id="OthersName" style="width: 280" value="<%=OthersName%>" maxlength="250">
        是否生效：<input name="ViewFlag" type="checkbox" value="1" <%if ViewFlag then response.write ("checked")%>> <font color="red">*</font> <input type="button" name="btn" value="复制标题" title="复制标题到：MetaDescription、MetaKeywords" onclick="CopyWebTitle(document.editForm.OthersName.value);"></td>
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
      <td align="right" class="forumRow">信息类别：</td>
      <td class="forumRowHighlight"><input name="SortID" type="text" id="SortID" style="width: 18; background-color:#fffff0" value="<%=SortID%>" readonly> <input name="SortPath" type="text" id="SortPath" style="width: 70; background-color:#fffff0" value="<%=SortPath%>" readonly> <input name="SortName" type="text" id="SortName" value="<%=SortName%>" style="width: 180; background-color:#fffff0" readonly> <a href="javaScript:OpenScript('SelectSort.Asp?Result=Others',500,500,'')">选择类别</a> <font color="red">*</font></td>
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
      <td align="right" class="forumRow">详细内容：</td>
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
Sub OthersEdit()
    Dim Action, rsRepeat, rs, sql
    Action = request.QueryString("Action")
    If Action = "SaveEdit" Then
        Set rs = server.CreateObject("adodb.recordset")
        If Len(Trim(request.Form("OthersName")))<3 Then
            response.Write ("<script language='javascript'>alert('请填写信息标题！');history.back(-1);</script>")
            response.End
        End If
        If Result = "Add" Then
            sql = "select * from Qianbo_Others"
            rs.Open sql, conn, 1, 3
            rs.addnew
            rs("OthersName") = Trim(Request.Form("OthersName"))
            If Request.Form("ViewFlag") = 1 Then
                rs("ViewFlag") = Request.Form("ViewFlag")
            Else
                rs("ViewFlag") = 0
            End If
            If Request.Form("SortID") = "" And Request.Form("SortPath") = "" Then
                response.Write ("<script language='javascript'>alert('请选择所属分类！');history.back(-1);</script>")
                response.End
            Else
                rs("SortID") = Request.Form("SortID")
                rs("SortPath") = Request.Form("SortPath")
            End If
            GroupIdName = Split(Request.Form("GroupID"), "┎╂┚")
            rs("GroupID") = GroupIdName(0)
            rs("Exclusive") = Trim(Request.Form("Exclusive"))
            rs("Content") = RTrim(Request.Form("Content"))
            rs("SeoKeywords") = Trim(Request.Form("SeoKeywords"))
            rs("SeoDescription") = Trim(Request.Form("SeoDescription"))
            rs("AddTime") = Now()
            rs("UpdateTime") = Now()
            rs.update
            rs.Close
            Set rs = Nothing
            Set rs = server.CreateObject("adodb.recordset")
            sql = "select top 1 ID from Qianbo_Others order by ID desc"
            rs.Open sql, conn, 1, 1
            ID = rs("ID")
            rs.Close
            Set rs = Nothing
            If ISHTML = 1 Then
                Call htmll("", "", ""&OtherName&""&Separated&""&ID&"."&HTMLName&"", "OtherView.Asp", "ID=", ID, "", "")
            End If
        End If
        If Result = "Modify" Then
            sql = "select * from Qianbo_Others where ID="&ID
            rs.Open sql, conn, 1, 3
            rs("OthersName") = Trim(Request.Form("OthersName"))
            If Request.Form("ViewFlag") = 1 Then
                rs("ViewFlag") = Request.Form("ViewFlag")
            Else
                rs("ViewFlag") = 0
            End If
            If Request.Form("SortID")<>"" And Request.Form("SortPath")<>"" Then
                rs("SortID") = Request.Form("SortID")
                rs("SortPath") = Request.Form("SortPath")
            Else
                response.Write ("<script language='javascript'>alert('请选择所属分类！');history.back(-1);</script>")
                response.End
            End If
            GroupIdName = Split(Request.Form("GroupID"), "┎╂┚")
            rs("GroupID") = GroupIdName(0)
            rs("Exclusive") = Trim(Request.Form("Exclusive"))
            rs("Content") = RTrim(Request.Form("Content"))
            rs("SeoKeywords") = Trim(Request.Form("SeoKeywords"))
            rs("SeoDescription") = Trim(Request.Form("SeoDescription"))
            rs("UpdateTime") = Now()
            rs.update
            rs.Close
            Set rs = Nothing
            If ISHTML = 1 Then
                Call htmll("", "", ""&OtherName&""&Separated&""&ID&"."&HTMLName&"", "OtherView.Asp", "ID=", ID, "", "")
            End If
        End If
        If ISHTML = 1 Then
            response.Write "<script language='javascript'>alert('设置成功，相关静态页面已更新！');location.replace('OthersList.Asp');</script>"
        Else
            response.Write "<script language='javascript'>alert('设置成功！');location.replace('OthersList.Asp');</script>"
        End If
    Else
        If Result = "Modify" Then
            Set rs = server.CreateObject("adodb.recordset")
            sql = "select * from Qianbo_Others where ID="& ID
            rs.Open sql, conn, 1, 1
            If rs.bof And rs.EOF Then
                response.Write ("数据库读取记录出错！")
                response.End
            End If
            OthersName = rs("OthersName")
            ViewFlag = rs("ViewFlag")
            SortName = SortText(rs("SortID"))
            SortID = rs("SortID")
            SortPath = rs("SortPath")
            GroupID = rs("GroupID")
            Exclusive = rs("Exclusive")
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
    sql = "Select * From Qianbo_OthersSort where ID="&ID
    rs.Open sql, conn, 1, 1
    SortText = rs("SortName")
    rs.Close
    Set rs = Nothing
End Function
%>