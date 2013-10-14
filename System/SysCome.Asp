<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/Version.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
If Request("Act") = "ClsCache" Then
    Call ClsCache()
End If

m_SQL = "select count(*) from Qianbo_Admin"
Set rs = conn.Execute(m_SQL)
m_ManageNumber = rs(0)
rs.Close
Set rs = Nothing

m_SQL = "select count(*) from Qianbo_Members"
Set rs = conn.Execute(m_SQL)
m_UserNumber = rs(0)
rs.Close
Set rs = Nothing

m_SQL = "select count(*) from Qianbo_Message"
Set rs = conn.Execute(m_SQL)
m_Message = rs(0)
rs.Close
Set rs = Nothing

m_SQL = "select count(*) from Qianbo_Message where ViewFlag = 1"
Set rs = conn.Execute(m_SQL)
m_MessageViewFlag = rs(0)
rs.Close
Set rs = Nothing

m_SQL = "select count(*) from Qianbo_About"
Set rs = conn.Execute(m_SQL)
m_About = rs(0)
rs.Close
Set rs = Nothing

m_SQL = "select count(*) from Qianbo_Download"
Set rs = conn.Execute(m_SQL)
m_Download = rs(0)
rs.Close
Set rs = Nothing

m_SQL = "select count(*) from Qianbo_FriendLink"
Set rs = conn.Execute(m_SQL)
m_FriendLink = rs(0)
rs.Close
Set rs = Nothing

m_SQL = "select count(*) from Qianbo_Jobs"
Set rs = conn.Execute(m_SQL)
m_Jobs = rs(0)
rs.Close
Set rs = Nothing

m_SQL = "select count(*) from Qianbo_News"
Set rs = conn.Execute(m_SQL)
m_News = rs(0)
rs.Close
Set rs = Nothing

m_SQL = "select count(*) from Qianbo_Order"
Set rs = conn.Execute(m_SQL)
m_Order = rs(0)
rs.Close
Set rs = Nothing

m_SQL = "select count(*) from Qianbo_Others"
Set rs = conn.Execute(m_SQL)
m_Others = rs(0)
rs.Close
Set rs = Nothing

m_SQL = "select count(*) from Qianbo_Products"
Set rs = conn.Execute(m_SQL)
m_Products = rs(0)
rs.Close
Set rs = Nothing

m_SQL = "select count(*) from Qianbo_Talents"
Set rs = conn.Execute(m_SQL)
m_Talents = rs(0)
rs.Close
Set rs = Nothing

Function IsObjInstalled(strClassString)
    On Error Resume Next
    IsObjInstalled = False
    Err = 0
    Dim xTestObj
    Set xTestObj = CreateObject(strClassString)
    If Err.Number = 0 Then IsObjInstalled = True
    Set xTestObj = Nothing
    Err = 0
End Function

Sub ShowObjectInstalled(strObjName)
    If IsObjInstalled(strObjName) Then
        response.Write "<font color=""red""><b>√</font>"
    Else
        response.Write "<font color=""red""><b>×</b></font>"
    End If
End Sub
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="Images/Admin_style.css">
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <th height="22" colspan="2" sytle="line-height:150%">【系统信息】</th>
  </tr>
  <tr>
    <td width="47%" class="forumRow">服务器类型：<%=Request.ServerVariables("OS")%>(IP：<%=Request.ServerVariables("local_addr")%>)</td>
    <td width="53%" class="forumRowHighlight">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr>
    <td class="forumRow">站点物理路径：<%=request.ServerVariables("APPL_PHYSICAL_PAth")%></td>
    <td class="forumRowHighlight">CDONTS组件：<%
On Error Resume Next
Server.CreateObject("CDONTS.NewMail")
If Err = 0 Then
    response.Write("<b><font color=""red"">√</font></b>")
Else
    response.Write("<b><font color=""red"">×</font></b>")
End If
Err = 0
%>
      Jmail邮箱组件：<%
If Not IsObjInstalled(theInstalledObjects(13)) Then
    response.Write("<b><font color=""red"">×</font></b>")
Else
    response.Write("<b><font color=""red"">√</font></b>")
End If
%></td>
  </tr>
  <tr>
    <td class="forumRow">FSO文本读写：<%ShowObjectInstalled("scripting.filesystemobject")%></td>
    <td class="forumRowHighlight">脚本超时时间：<%=Server.ScriptTimeout%>秒</td>
  </tr>
  <tr>
    <td class="forumRow">客户端操作系统：<%
Dim thesoft, vOS
thesoft = Request.ServerVariables("HTTP_USER_AGENT")
If InStr(thesoft, "Windows NT 5.0") Then
    vOS = "Microsoft Windows 2000"
ElseIf InStr(thesoft, "Windows NT 5.2") Then
    vOs = "Microsoft Windows 2003"
ElseIf InStr(thesoft, "Windows NT 5.1") Then
    vOs = "Microsoft Windows XP"
ElseIf InStr(thesoft, "Windows NT") Then
    vOs = "Microsoft Windows NT"
ElseIf InStr(thesoft, "Windows 9") Then
    vOs = "Microsoft Windows 9x"
ElseIf InStr(thesoft, "unix") Or InStr(thesoft, "linux") Or InStr(thesoft, "SunOS") Or InStr(thesoft, "BSD") Then
    vOs = "类Unix"
ElseIf InStr(thesoft, "Mac") Then
    vOs = "Mac"
Else
    vOs = "Other"
End If
response.Write(vOs)
%></td>
    <td class="forumRowHighlight">返回服务器处理请求的端口：<%=Request.ServerVariables("SERVER_PORT")%></td>
  </tr>
</table>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <th height="22" colspan="2" sytle="line-height:150%">【版本信息】</th>
  </tr>
  <tr>
    <td width="47%" class="forumRow">网联动力营销策划有限公司 <%=Str_Soft_Version%></td>
    <td width="53%" class="forumRowHighlight">程序开发：千博网络技术部</td>
  </tr>
  <tr>
    <td width="47%" class="forumRow">管 理 员：<%=m_ManageNumber%>个 注册会员：<%=m_UserNumber%>个 留言：<%=m_Message%>(已审<%=m_MessageViewFlag%>条) 应聘信息：<%=m_Talents%>条</td>
    <td width="53%" class="forumRowHighlight">企业信息：<%=m_About%>条 下载信息：<%=m_Download%>条 友情链接：<%=m_FriendLink%>条 人才信息：<%=m_Jobs%>条</td>
  </tr>
  <tr>
    <td width="47%" class="forumRow">新闻动态：<%=m_News%>条 在线订单：<%=m_Order%>条</td>
    <td width="53%" class="forumRowHighlight">其他信息：<%=m_Others%>条 公司产品：<%=m_Products%>条</td>
  </tr>
</table>
<br />
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center">Copyright (c) 2010 <a href="http://www.Qianbo.com.cn" title="千博网络：创造价值，成就你我！" target="_blank"><font face="Verdana, arial, helvetica, sans-serif" size="1"><b>www.Qianbo<font color="#CC0000">.com.cn</font></b></font></a> All Rights Reserved.</td>
  </tr>
</table>
<br />