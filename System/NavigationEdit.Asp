<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
If InStr(CompanyAdminPurview, "|2,") = 0 Then
    response.Write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
    response.End
End If
Dim Result
Result = request.QueryString("Result")
Dim ID, NavName, ViewFlag, NavUrl, HtmlNavUrl, OutFlag, Remark
ID = request.QueryString("ID")
Call NavEdit()
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
<form name="editForm" method="post" action="NavigationEdit.Asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <th height="22" colspan="2" sytle="line-height:150%">【<%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>导航】</th>
  </tr>
  <tr>
    <td width="20%" align="right" class="forumRow">导航名称：</td>
    <td width="80%" class="forumRowHighlight"><input name="NavName" type="text" id="NavName" style="width: 200" value="<%=NavName%>" maxlength="100"> 发布：<input name="ViewFlag" type="checkbox" value="1" <%if ViewFlag then response.write ("checked")%>> <font color="red">*</font></td>
  </tr>
  <tr>
    <td align="right" class="forumRow">动态页链接网址：</td>
    <td class="forumRowHighlight"><input name="NavUrl" type="text" id="NavUrl" style="width: 500" value="<%=NavUrl%>"> <font color="red">*</font></td>
  </tr>
  <tr>
    <td align="right" class="forumRow">静态页链接网址：</td>
    <td class="forumRowHighlight"><input name="HtmlNavUrl" type="text" id="HtmlNavUrl" style="width: 500" value="<%=HtmlNavUrl%>"> <font color="red">*</font></td>
  </tr>
  <tr>
    <td width="20%" align="right" class="forumRow">链接状态：</td>
    <td width="80%" class="forumRowHighlight"><input name="OutFlag" type="checkbox" value="1" <%if OutFlag then response.write ("checked")%>>是否外部链接</td>
  </tr>
  <tr>
    <td align="right" class="forumRow">备注：</td>
    <td class="forumRowHighlight"><textarea name="Remark" rows="8" id="Remark" style="width: 500"><%=Remark%></textarea></td>
  </tr>
  <tr>
    <td width="20%" align="right" class="forumRow"></td>
    <td width="80%" class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="保存"> <input type="button" value="返回上一页" onclick="history.back(-1)"></td>
  </tr>
  </form>
</table>
<%
Sub NavEdit()
    Dim Action, rsCheckAdd, rs, sql
    Action = request.QueryString("Action")
    If Action = "SaveEdit" Then
        Set rs = server.CreateObject("adodb.recordset")
        If Len(Trim(request.Form("NavName")))<2 Then
            response.Write ("<script language='javascript'>alert('请填写导航名称并保持至少在一个汉字以上！');history.back(-1);</script>")
            response.End
        End If
        If Trim(Request.Form("NavName")) = "" Or Trim(Request.Form("NavUrl")) = "" Or Trim(Request.Form("HtmlNavUrl")) = "" Then
            response.Write ("<script language='javascript'>alert('请填写导航名称和链接网址！');history.back(-1);</script>")
            response.End
        End If
        If Result = "Add" Then
            sql = "select * from Qianbo_Navigation"
            rs.Open sql, conn, 1, 3
            rs.addnew
            rs("NavName") = Trim(Request.Form("NavName"))
            rs("NavUrl") = Trim(Request.Form("NavUrl"))
            rs("HtmlNavUrl") = Trim(Request.Form("HtmlNavUrl"))
            If Request.Form("ViewFlag") = 1 Then
                rs("ViewFlag") = Request.Form("ViewFlag")
            Else
                rs("ViewFlag") = 0
            End If
            If Request.Form("OutFlag") = 1 Then
                rs("OutFlag") = Request.Form("OutFlag")
            Else
                rs("OutFlag") = 0
            End If
            rs("Remark") = Trim(Request.Form("Remark"))
            rs("Sequence") = 99
            rs("AddTime") = Now()
        End If
        If Result = "Modify" Then
            sql = "select * from Qianbo_Navigation where ID="&ID
            rs.Open sql, conn, 1, 3
            rs("NavName") = Trim(Request.Form("NavName"))
            rs("NavUrl") = Trim(Request.Form("NavUrl"))
            rs("HtmlNavUrl") = Trim(Request.Form("HtmlNavUrl"))
            If Request.Form("ViewFlag") = 1 Then
                rs("ViewFlag") = Request.Form("ViewFlag")
            Else
                rs("ViewFlag") = 0
            End If
            If Request.Form("OutFlag") = 1 Then
                rs("OutFlag") = Request.Form("OutFlag")
            Else
                rs("OutFlag") = 0
            End If
            rs("Remark") = Trim(Request.Form("Remark"))
        End If
        rs.update
        rs.Close
        Set rs = Nothing
        response.Write "<script language='javascript'>alert('设置成功！');location.replace('NavigationList.Asp');</script>"
    Else
        If Result = "Modify" Then
            Set rs = server.CreateObject("adodb.recordset")
            sql = "select * from Qianbo_Navigation where ID="& ID
            rs.Open sql, conn, 1, 1
            NavName = rs("NavName")
            ViewFlag = rs("ViewFlag")
            Remark = rs("Remark")
            OutFlag = rs("OutFlag")
            NavUrl = rs("NavUrl")
            HtmlNavUrl = rs("HtmlNavUrl")
            rs.Close
            Set rs = Nothing
        End If
    End If
End Sub
%>