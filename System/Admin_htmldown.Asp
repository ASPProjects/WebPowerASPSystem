﻿<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="Admin_htmlconfig.Asp"-->
<link rel="stylesheet" href="Images/Admin_style.css">
<%
If InStr(CompanyAdminPurview, "|34,") = 0 Then
    response.Write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
    response.End
End If
If ISHTML = 0 Then
    response.Write "<script language='javascript'>alert('请先在【系统参数配置】中将静态HTML设置为开启！');history.go(-1);</script>"
    response.End
End If
%>
<br />
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="300"><table width="100%" border="0" cellspacing="0" cellpadding="1">
        <tr>
          <td style="border-bottom: #ccc 1px solid; border-top: #ccc 1px solid; border-left: #ccc 1px solid; border-right: #ccc 1px solid"><img src="Images/Survey_1.gif" width="0" height="16" id="bar_img" name="bar_img" align="absmiddle"></td>
        </tr>
      </table></td>
    <td><span id="bar_txt2" name="bar_txt2" style="font-size:12px; color:red;"></span><span id="bar_txt1" name="bar_txt1" style="font-size:12px">0</span><span style="font-size:12px">%</span></td>
  </tr>
</table>
<%
Call HtmlDown
conn.Close
Set conn = Nothing
%>