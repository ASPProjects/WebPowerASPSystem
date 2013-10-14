<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/Md5.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<%
If InStr(CompanyAdminPurview, "|31,") = 0 Then
    response.Write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
    response.End
End If
Dim Result
Result = request.QueryString("Result")
Dim ID, mMemName, mRealName, mPassword, vPassword, mSex, mGroupID, mGroupName, mGroupIdName
Dim mCompany, mAddress, mZipCode, mTelephone, mFax, mMobile, mEmail, mHomepage, mWorking
ID = request.QueryString("ID")
Call MemEdit()
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="editForm" method="post" action="MemEdit.Asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">【<%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>会员信息】</th>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">登录名称：</td>
      <td width="80%" class="forumRowHighlight"><input name="MemName" type="text" id="MemName" style="width: 120" value="<%=mMemName%>" maxlength="16" <%if Result="Modify" then response.write ("readonly")%>> <font color="red">*</font>3-16个字符</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">真实姓名：</td>
      <td class="forumRowHighlight"><input name="RealName" type="text" id="RealName" style="width: 120" value="<%=mRealName%>" maxlength="16"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">登录密码：</td>
      <td class="forumRowHighlight"><input name="Password" type="password" id="Password" maxlength="20" style="width: 120"> <font color="red">*</font>6-16个字符</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">确认密码：</td>
      <td class="forumRowHighlight"><input name="vPassword" type="password" id="vPassword" maxlength="20" style="width: 120"> <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">性别：</td>
      <td class="forumRowHighlight"><input type="radio" name="sex" value="先生" <%if mSex="先生" then response.write ("checked")%>>先生 <input type="radio" name="sex" value="女士" <%if mSex="女士" then response.write ("checked")%>>女士</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">会员组别：</td>
      <td class="forumRowHighlight"><select name="GroupID"><% call SelectGroup() %></select></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">公司名称：</td>
      <td class="forumRowHighlight"><input name="Company" type="text" id="Company" style="width: 280" value="<%=mCompany%>" maxlength="100"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">地址：</td>
      <td class="forumRowHighlight"><input name="Address" type="text" id="Address" style="width: 280" value="<%=mAddress%>" maxlength="100"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">邮政编码：</td>
      <td class="forumRowHighlight"><input name="ZipCode" type="text" id="ZipCode" style="width: 120" value="<%=mZipCode%>" maxlength="16"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">联系电话：</td>
      <td class="forumRowHighlight"><input name="Telephone" type="text" id="Telephone" style="width: 120" value="<%=mTelephone%>" maxlength="50"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">传真号码：</td>
      <td class="forumRowHighlight"><input name="Fax" type="text" id="Fax" style="width: 120" value="<%=mFax%>" maxlength="16"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">手机号码：</td>
      <td class="forumRowHighlight"><input name="Mobile" type="text" id="Mobile" style="width: 120" value="<%=mMobile%>" maxlength="16"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">电子信箱：</td>
      <td class="forumRowHighlight"><input name="Email" type="text" id="Email" style="width: 280" value="<%=mEmail%>" maxlength="50"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">网址：</td>
      <td class="forumRowHighlight"><input name="HomePage" type="text" id="HomePage" style="width: 280" value="<%=mHomePage%>" maxlength="50"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">生效：</td>
      <td class="forumRowHighlight"><input name="Working" type="checkbox"  value="1" <%if mWorking then response.write ("checked")%>></td>
    </tr>
    <tr>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="保存"> <input type="button" value="返回上一页" onclick="history.back(-1)"></td>
    </tr>
  </form>
</table>
<%
Sub MemEdit()
    Dim Action, rs, sql
    Action = request.QueryString("Action")
    If Action = "SaveEdit" Then
        Set rs = server.CreateObject("adodb.recordset")
        If Result = "Add" Then
            If Not conn.Execute("select MemName from Qianbo_Members where MemName='" & Trim(Request.Form("MemName")) & "'").EOF Then
                response.Write "<script language='javascript'>alert('" & Trim(Request.Form("MemName")) & "会员名已经存在！');history.back(-1);</script>"
                response.End
            End If
            sql = "select * from Qianbo_Members"
            rs.Open sql, conn, 1, 3
            rs.addnew
            rs("MemName") = Trim(Request.Form("MemName"))
            rs("RealName") = StrReplace(Trim(Request.Form("RealName")))
            If Len(Trim(Request.Form("Password")))<6 Or Len(Trim(Request.Form("Password")))>16 Then
                response.Write "<script language='javascript'>alert('请填写登录密码(6-16个字符)！');history.back(-1);</script>"
                response.End
            End If
            If Request.Form("Password")<>Request.Form("vPassword") Then
                response.Write "<script language='javascript'>alert('两次输入的密码不同！');history.back(-1);</script>"
                response.End
            End If
            rs("Password") = Md5(Request.Form("Password"))
            rs("Sex") = Request.Form("Sex")
            mGroupIdName = Split(Request.Form("GroupID"), "┎╂┚")
            rs("GroupID") = mGroupIdName(0)
            rs("GroupName") = mGroupIdName(1)
            rs("Company") = StrReplace(Trim(Request.Form("Company")))
            rs("Address") = StrReplace(Trim(Request.Form("Address")))
            rs("ZipCode") = StrReplace(Trim(Request.Form("ZipCode")))
            rs("Telephone") = StrReplace(Trim(Request.Form("Telephone")))
            rs("Fax") = StrReplace(Trim(Request.Form("Fax")))
            rs("Mobile") = StrReplace(Trim(Request.Form("Mobile")))
            rs("Email") = Trim(Request.Form("Email"))
            rs("HomePage") = StrReplace(Trim(Request.Form("HomePage")))
            If Request.Form("Working") = 1 Then
                rs("Working") = Request.Form("Working")
            Else
                rs("Working") = 0
            End If
            rs("AddTime") = Now()
        End If
        If Result = "Modify" Then
            sql = "select * from Qianbo_Members where ID="&ID
            rs.Open sql, conn, 1, 3
            rs("MemName") = Trim(Request.Form("MemName"))
            rs("RealName") = StrReplace(Trim(Request.Form("RealName")))
            If Trim(Request.Form("Password"))<>"" Then
                If Len(Trim(Request.Form("Password")))<6 Or Len(Trim(Request.Form("Password")))>16 Then
                    response.Write "<script language='javascript'>alert('请填写登录密码(6-16个字符)！');history.back(-1);</script>"
                    response.End
                End If
                If Request.Form("Password")<>Request.Form("vPassword") Then
                    response.Write "<script language='javascript'>alert('两次输入的密码不同！');history.back(-1);</script>"
                    response.End
                End If
                rs("Password") = Md5(Request.Form("Password"))
            End If
            rs("Sex") = Request.Form("Sex")
            mGroupIdName = Split(Request.Form("GroupID"), "┎╂┚")
            rs("GroupID") = mGroupIdName(0)
            rs("GroupName") = mGroupIdName(1)
            rs("Company") = StrReplace(Trim(Request.Form("Company")))
            rs("Address") = StrReplace(Trim(Request.Form("Address")))
            rs("ZipCode") = StrReplace(Trim(Request.Form("ZipCode")))
            rs("Telephone") = StrReplace(Trim(Request.Form("Telephone")))
            rs("Fax") = StrReplace(Trim(Request.Form("Fax")))
            rs("Mobile") = StrReplace(Trim(Request.Form("Mobile")))
            rs("Email") = StrReplace(Trim(Request.Form("Email")))
            rs("HomePage") = StrReplace(Trim(Request.Form("HomePage")))
            If Request.Form("Working") = 1 Then
                rs("Working") = Request.Form("Working")
            Else
                rs("Working") = 0
            End If
        End If
        rs.update
        rs.Close
        Set rs = Nothing
        response.Write "<script language='javascript'>alert('设置成功！');location.replace('MemList.Asp');</script>"
    Else
        If Result = "Modify" Then
            Set rs = server.CreateObject("adodb.recordset")
            sql = "select * from Qianbo_Members where ID="& ID
            rs.Open sql, conn, 1, 1
            mMemName = rs("MemName")
            mRealName = rs("RealName")
            mSex = rs("Sex")
            mGroupID = rs("GroupID")
            mGroupName = rs("GroupName")
            mCompany = rs("Company")
            mAddress = rs("Address")
            mZipCode = rs("ZipCode")
            mTelephone = rs("Telephone")
            mFax = rs("Fax")
            mMobile = rs("Mobile")
            mEmail = rs("Email")
            mHomepage = rs("Homepage")
            mWorking = rs("Working")
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
        If mGroupID = rs("GroupID") Then response.Write ("selected")
        response.Write(">"&rs("GroupName")&"</option>")
        rs.movenext
    Wend
    rs.Close
    Set rs = Nothing
End Sub
%>