<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<%
If InStr(CompanyAdminPurview, "|27,") = 0 Then
    response.Write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
    response.End
End If
Dim Result
Result = request.QueryString("Result")
Dim ReplyContent, ReplyTime, ID, JobID, TalentsName
Dim mLinkman, mBirthDate, mStature, mMarriage, mRegResidence, mEduResume, mJobResume, mAddress, mZipCode, mTelephone, mMobile, mEmail, AddTime
ID = request.QueryString("ID")
Call TalentsEdit()
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="editForm" method="post" action="TalentsEdit.Asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">【<%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>人才】</th>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">应聘职位：</td>
      <td width="80%" class="forumRowHighlight"><input name="TalentsName" type="text" id="TalentsName" style="width: 280" value="<%=TalentsName%>" readonly> <a href="JobsEdit.Asp?Result=Modify&ID=<%=JobID%>" target="main">查看招聘</a></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">应聘人：</td>
      <td class="forumRowHighlight"><%=mLinkman%></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">出生日期：</td>
      <td class="forumRowHighlight"><input name="BirthDate" type="text" id="BirthDate" style="width: 180" value="<%=mBirthDate%>" readonly></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">身高：</td>
      <td class="forumRowHighlight"><input name="Stature" type="text" id="Stature" style="width: 180" value="<%=mStature%>" readonly></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">婚姻状况：</td>
      <td class="forumRowHighlight"><input name="Marriage" type="text" id="Marriage" style="width: 180" value="<%=mMarriage%>" readonly></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">户口所在地：</td>
      <td class="forumRowHighlight"><input name="RegResidence" type="text" id="RegResidence" style="width: 280" value="<%=mRegResidence%>" readonly></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">教育经历：</td>
      <td class="forumRowHighlight"><textarea name="EduResume" rows="10" id="EduResume" style="width: 500" readonly><%=mEduResume%></textarea></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">工作经历：</td>
      <td class="forumRowHighlight"><textarea name="JobResume" rows="10" id="JobResume" style="width: 500" readonly><%=mJobResume%></textarea></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">通信地址：</td>
      <td class="forumRowHighlight"><input name="Address" type="text" id="Address" style="width: 280" value="<%=mAddress%>" readonly></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">邮政编码：</td>
      <td class="forumRowHighlight"><input name="ZipCode" type="text" id="ZipCode" style="width: 80" value="<%=mZipCode%>" readonly></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">联系电话：</td>
      <td class="forumRowHighlight"><input name="Telephone" type="text" id="Telephone" style="width: 180" value="<%=mTelephone%>" readonly></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">手机号码：</td>
      <td class="forumRowHighlight"><input name="Mobile" type="text" id="Mobile" style="width: 180" value="<%=mMobile%>" readonly></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">电子信箱：</td>
      <td class="forumRowHighlight"><input name="Email" type="text" id="Email" style="width: 180" value="<%=mEmail%>" readonly></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">提交时间：</td>
      <td class="forumRowHighlight"><input name="AddTime" type="text" id="AddTime" style="width: 180" value="<%=AddTime%>" readonly></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">回复时间：</td>
      <td class="forumRowHighlight"><input name="ReplyTime" type="text" id="ReplyTime" style="width: 180" value="<%=ReplyTime%>" readonly></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">回复内容：</td>
      <td class="forumRowHighlight"><textarea name="ReplyContent" rows="6" id="ReplyContent" style="width: 500"><%=ReplyContent%></textarea></td>
    </tr>
    <tr>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="保存"> <input type="button" value="返回上一页" onclick="history.back(-1)"></td>
    </tr>
  </form>
</table>
<%
Sub TalentsEdit()
    Dim Action, rsCheckAdd, rs, sql
    Action = request.QueryString("Action")
    If Action = "SaveEdit" Then
        Set rs = server.CreateObject("adodb.recordset")
        If Result = "Modify" Then
            sql = "select * from Qianbo_Talents where ID="&ID
            rs.Open sql, conn, 1, 3
            rs("ReplyContent") = StrReplace(Request.Form("ReplyContent"))
            If Not (Trim(request.Form("ReplyContent")) = "" Or Trim(request.Form("ReplyTime"))<>"") Then
                rs("ReplyTime") = Now()
            End If
        End If
        rs.update
        rs.Close
        Set rs = Nothing
        response.Write "<script language='javascript'>alert('编辑、回复人才信息成功！');location.replace('TalentsList.Asp');</script>"
    Else
        If Result = "Modify" Then
            Set rs = server.CreateObject("adodb.recordset")
            sql = "select * from Qianbo_Talents where ID="& ID
            rs.Open sql, conn, 1, 1
            JobID = rs("JobID")
            TalentsName = rs("TalentsName")
            mLinkman = GuestInfo(rs("MemID"), rs("Linkman"), rs("Sex"))
            mBirthDate = rs("BirthDate")
            mStature = rs("Stature")
            mMarriage = rs("Marriage")
            mRegResidence = rs("RegResidence")
            mEduResume = ReStrReplace(rs("EduResume"))
            mJobResume = ReStrReplace(rs("JobResume"))
            mAddress = rs("Address")
            mZipCode = rs("ZipCode")
            mTelephone = rs("Telephone")
            mMobile = rs("Mobile")
            mEmail = rs("Email")
            AddTime = rs("AddTime")
            ReplyContent = ReStrReplace(rs("ReplyContent"))
            ReplyTime = rs("ReplyTime")
            rs.Close
            Set rs = Nothing
        End If
    End If
End Sub

Function GuestInfo(ID, Guest, Sex)
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "Select * From Qianbo_Members where ID="&ID
    rs.Open sql, conn, 1, 1
    If rs.EOF Then
        GuestInfo = Guest & "&nbsp;" & Sex
    Else
        GuestInfo = "<font color='green'>会员：</font><a href='MemEdit.Asp?Result=Modify&ID="&ID&"'>"&Guest&"</a>"&Sex
    End If
    rs.Close
    Set rs = Nothing
End Function
%>