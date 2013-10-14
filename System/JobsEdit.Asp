<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="Admin_htmlconfig.Asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<%
If InStr(CompanyAdminPurview, "|18,") = 0 Then
    response.Write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
    response.End
End If
Dim Result
Result = request.QueryString("Result")
Dim ID, JobName, ViewFlag, JobAddress, JobNumber, Emolument, StartDate, EndDate, Responsibility, Requirement
Dim eEmployer, eContact, eTel, eAddress, ePostCode, eEmail
ID = request.QueryString("ID")
Call JobsEdit()
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="editForm" method="post" action="JobsEdit.Asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">【<%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>招聘】</th>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">职位名称：</td>
      <td width="80%" class="forumRowHighlight"><input name="JobName" type="text" id="JobName" style="width: 280" value="<%=JobName%>" maxlength="280">
        是否生效：<input name="ViewFlag" type="checkbox" value="1" <%if ViewFlag then response.write ("checked")%>> <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">工作地点：</td>
      <td class="forumRowHighlight"><input name="JobAddress" type="text" style="width: 180" value="<%=JobAddress%>" maxlength="180"> <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">招聘人数：</td>
      <td class="forumRowHighlight"><input name="JobNumber" type="text" style="width: 80" value="<%=JobNumber%>" maxlength="80"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">薪水待遇：</td>
      <td class="forumRowHighlight"><input name="Emolument" type="text" style="width: 80" value="<%=Emolument%>" maxlength="80"> <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">时间：</td>
      <td class="forumRowHighlight"><input name="StartDate" type="text" id="StartDate" style="width: 120" value="<% if StartDate="" then response.write now() else response.write (StartDate) end if%>" maxlength="18"> → <input name="EndDate" type="text" id="EndDate" style="width: 120" value="<% if EndDate="" then response.write (DateAdd("m",3,now())) else response.write (EndDate) end if%>" maxlength="18"> <font color="red">*</font> 默认为当前时间开始，三个月后结束(可手工修改)。</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">工作职责：</td>
      <td class="forumRowHighlight"><input name="Responsibility" type="text" style="width: 500" value="<%=Responsibility%>" maxlength="500"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">职位要求：</td>
      <td class="forumRowHighlight"><input name="Requirement" type="text" style="width: 500" value="<%=Requirement%>" maxlength="500"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">招聘单位：</td>
      <td class="forumRowHighlight"><input name="eEmployer" type="text" style="width: 280" value="<%=eEmployer%>" maxlength="280"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">联系人：</td>
      <td class="forumRowHighlight"><input name="eContact" type="text" style="width: 180" value="<%=eContact%>" maxlength="180"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">联系电话：</td>
      <td class="forumRowHighlight"><input name="eTel" type="text" style="width: 180" value="<%=eTel%>" maxlength="180"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">联系地址：</td>
      <td class="forumRowHighlight"><input name="eAddress" type="text" style="width: 280" value="<%=eAddress%>" maxlength="280"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">邮政编码：</td>
      <td class="forumRowHighlight"><input name="ePostCode" type="text" style="width: 80" value="<%=ePostCode%>" maxlength="80"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">电子信箱：</td>
      <td class="forumRowHighlight"><input name="eEmail" type="text" style="width: 280" value="<%=eEmail%>" maxlength="280"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="保存"> <input type="button" value="返回上一页" onclick="history.back(-1)"></td>
    </tr>
  </form>
</table>
<%
Sub JobsEdit()
    Dim Action, rsCheckAdd, rs, sql
    Action = request.QueryString("Action")
    If Action = "SaveEdit" Then
        Set rs = server.CreateObject("adodb.recordset")
        If Len(Trim(request.Form("JobName")))<1 Then
            response.Write ("<script language='javascript'>alert('请填写招聘职位名称！');history.back(-1);</script>")
            response.End
        End If
        If Len(Trim(request.Form("JobAddress")))<1 Or Len(Trim(request.Form("Emolument")))<1 Then
            response.Write ("<script language='javascript'>alert('请填写工作地点、薪水待遇！');history.back(-1);</script>")
            response.End
        End If
        If Not IsNumeric(Trim(request.Form("JobNumber"))) Then
            response.Write ("<script language='javascript'>alert('请正确填写职位数量！');history.back(-1);</script>")
            response.End
        End If
        If Not (IsDate(Trim(request.Form("StartDate"))) Or IsDate(Trim(request.Form("EndDate")))) Then
            response.Write ("<script language='javascript'>alert('请正确填写开始、结束日期！');history.back(-1);</script>")
            response.End
        End If
        If Result = "Add" Then
            sql = "select * from Qianbo_Jobs"
            rs.Open sql, conn, 1, 3
            rs.addnew
            rs("JobName") = Trim(Request.Form("JobName"))
            If Request.Form("ViewFlag") = 1 Then
                rs("ViewFlag") = Request.Form("ViewFlag")
            Else
                rs("ViewFlag") = 0
            End If
            rs("JobAddress") = Trim(Request.Form("JobAddress"))
            rs("JobNumber") = Trim(Request.Form("JobNumber"))
            rs("Emolument") = Trim(Request.Form("Emolument"))
            rs("StartDate") = Trim(Request.Form("StartDate"))
            rs("EndDate") = Trim(Request.Form("EndDate"))
            rs("Responsibility") = StrReplace(Request.Form("Responsibility"))
            rs("Requirement") = StrReplace(Request.Form("Requirement"))
            rs("AddTime") = Now()
            rs("UpdateTime") = rs("AddTime")
            rs("eEmployer") = Trim(Request.Form("eEmployer"))
            rs("eContact") = Trim(Request.Form("eContact"))
            rs("eTel") = Trim(Request.Form("eTel"))
            rs("eAddress") = Trim(Request.Form("eAddress"))
            rs("ePostcode") = Trim(Request.Form("ePostcode"))
            rs("eEmail") = Trim(Request.Form("eEmail"))
            rs.update
            rs.Close
            Set rs = Nothing
            Set rs = server.CreateObject("adodb.recordset")
            sql = "select top 1 ID from Qianbo_Jobs order by ID desc"
            rs.Open sql, conn, 1, 1
            ID = rs("ID")
            rs.Close
            Set rs = Nothing
            If ISHTML = 1 Then
                Call htmll("", "", ""&JobNameDiy&""&Separated&""&ID&"."&HTMLName&"", "JobsView.Asp", "ID=", ID, "", "")
            End If
        End If
        If Result = "Modify" Then
            sql = "select * from Qianbo_Jobs where ID="&ID
            rs.Open sql, conn, 1, 3
            rs("JobName") = Trim(Request.Form("JobName"))
            If Request.Form("ViewFlag") = 1 Then
                rs("ViewFlag") = Request.Form("ViewFlag")
            Else
                rs("ViewFlag") = 0
            End If
            rs("JobAddress") = Trim(Request.Form("JobAddress"))
            rs("JobNumber") = Trim(Request.Form("JobNumber"))
            rs("Emolument") = Trim(Request.Form("Emolument"))
            rs("StartDate") = Trim(Request.Form("StartDate"))
            rs("EndDate") = Trim(Request.Form("EndDate"))
            rs("Responsibility") = StrReplace(Request.Form("Responsibility"))
            rs("Requirement") = StrReplace(Request.Form("Requirement"))
            rs("UpdateTime") = Now()
            rs("eEmployer") = Trim(Request.Form("eEmployer"))
            rs("eContact") = Trim(Request.Form("eContact"))
            rs("eTel") = Trim(Request.Form("eTel"))
            rs("eAddress") = Trim(Request.Form("eAddress"))
            rs("ePostcode") = Trim(Request.Form("ePostcode"))
            rs("eEmail") = Trim(Request.Form("eEmail"))
            rs.update
            rs.Close
            Set rs = Nothing
            If ISHTML = 1 Then
                Call htmll("", "", ""&JobNameDiy&""&Separated&""&ID&"."&HTMLName&"", "JobsView.Asp", "ID=", ID, "", "")
            End If
        End If
        If ISHTML = 1 Then
            response.Write "<script language='javascript'>alert('设置成功，相关静态页面已更新！');location.replace('JobsList.Asp');</script>"
        Else
            response.Write "<script language='javascript'>alert('设置成功！');location.replace('JobsList.Asp');</script>"
        End If
    Else
        If Result = "Modify" Then
            Set rs = server.CreateObject("adodb.recordset")
            sql = "select * from Qianbo_Jobs where ID="& ID
            rs.Open sql, conn, 1, 1
            JobName = rs("JobName")
            ViewFlag = rs("ViewFlag")
            JobAddress = rs("JobAddress")
            JobNumber = rs("JobNumber")
            Emolument = rs("Emolument")
            StartDate = rs("StartDate")
            EndDate = rs("EndDate")
            Responsibility = ReStrReplace(rs("Responsibility"))
            Requirement = ReStrReplace(rs("Requirement"))
            eEmployer = rs("eEmployer")
            eContact = rs("eContact")
            eTel = rs("eTel")
            eAddress = rs("eAddress")
            ePostcode = rs("ePostcode")
            eEmail = rs("eEmail")
            rs.Close
            Set rs = Nothing
        End If
    End If
End Sub
%>