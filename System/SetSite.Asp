<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/Version.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<!--#include file="Fckeditor/Fckeditor.Asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript">
<!--
function setpic(){
    var arr = showModalDialog("eWebEditor/customDialog/img.htm", "", "dialogWidth:30em; dialogHeight:26em; status:0;help=no");
    if (arr ==null){
        alert("系统提示：当前没有上传图片，界面预览图为空，用户可以重新上传图片！");
    }
    if (arr !=null){
        editForm.SiteLogo.value=arr;
    }
}
//-->
</script>
<%
If InStr(CompanyAdminPurview, "|1,") = 0 Then
    response.Write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
    response.End
End If
Select Case request.QueryString("Action")
    Case "Save"
        SaveSiteInfo
    Case "SaveConst"
        SaveConstInfo
    Case Else
        ViewSiteInfo
End Select
%>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="editForm" method="post" action="?Action=Save">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">【系统主参数设置】</th>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">网站标题：</td>
      <td width="80%" class="forumRowHighlight"><input name="SiteTitle" type="text" id="SiteTitle" style="width: 280" value="<%=SiteTitle%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">公司网址：</td>
      <td class="forumRowHighlight"><input name="SiteUrl" type="text" id="SiteUrl" style="width: 280" value="<%=SiteUrl%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">公司名称：</td>
      <td class="forumRowHighlight"><input name="ComName" type="text" id="ComName" style="width: 280" value="<%=ComName%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">详细地址：</td>
      <td class="forumRowHighlight"><input name="Address" type="text" id="Address" style="width: 280" value="<%=Address%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">邮政编码：</td>
      <td class="forumRowHighlight"><input name="ZipCode" type="text" id="ZipCode" style="width: 180" value="<%=ZipCode%>" maxlength="6">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">联系电话：</td>
      <td class="forumRowHighlight"><input name="Telephone" type="text" id="Telephone" style="width: 180" value="<%=Telephone%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">传真号码：</td>
      <td class="forumRowHighlight"><input name="Fax" type="text" id="Fax" style="width: 180" value="<%=Fax%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">电子邮箱：</td>
      <td class="forumRowHighlight"><input name="Email" type="text" id="Email" style="width: 180" value="<%=Email%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">Keywords优化：</td>
      <td class="forumRowHighlight"><textarea name="Keywords" rows="4"  id="Keywords" style="width: 500"><%=Keywords%></textarea></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">Description优化：</td>
      <td class="forumRowHighlight"><textarea name="Descriptions" rows="4" id="Descriptions" style="width: 500"><%=Descriptions%></textarea></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">ICP备案号：</td>
      <td class="forumRowHighlight"><input name="IcpNumber" type="text" id="IcpNumber" style="width: 180" value="<%=IcpNumber%>"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">首页视频地址：</td>
      <td class="forumRowHighlight"><input name="Video" type="text" id="Video" style="width: 500" value="<%=Video%>">
        <br />
        视频格式支持：Mp3/Avi/Wmv/Asf/Mov/Rm/Ra/Ram/Rmvb/Swf</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">访客留言设置：</td>
      <td class="forumRowHighlight"><input name="MesViewFlag" type="checkbox" id="MesViewFlag" value="1" <%if MesViewFlag then response.write ("checked")%>>
        无须审核</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">首页Logo设置：</td>
      <td class="forumRowHighlight"><input name="SiteLogo" type="text" id="SiteLogo" style="width: 500" value="<%=SiteLogo%>"> <input type="button" value="上传图片" onClick="setpic();"></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">首页公司简介：</td>
      <td class="forumRowHighlight"><%
Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath = "FCKeditor/"
oFCKeditor.ToolbarSet = "Default"
oFCKeditor.Width = "100%"
oFCKeditor.Height = "350"
oFCKeditor.Value = SiteDetail
oFCKeditor.Create "SiteDetail"
%></td>
    </tr>
    <tr>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="保存主参数设置"></td>
    </tr>
  </form>
</table>
<br />
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <form name="ConstForm" method="post" action="?Action=SaveConst">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">【附加参数设置】</th>
    </tr>
    <tr>
      <td width="20%" align="right" class="forumRow">系统安装目录：</td>
      <td width="80%" class="forumRowHighlight"><input name="SysRootDir" type="text" id="SysRootDir" style="width: 280" value="<%=SysRootDir%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">数据库目录：</td>
      <td class="forumRowHighlight"><input name="SiteDataPath" type="text" id="SiteDataPath" style="width: 280" value="<%=SiteDataPath%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">是否启用防注入系统：</td>
      <td class="forumRowHighlight"><input name="EnableStopInjection" type="radio" value="True" <%if EnableStopInjection = True then%> checked<%end if%>>
        启用
        <input name="EnableStopInjection" type="radio" value="False" <%if EnableStopInjection = False then%> checked<%end if%>>
        不启用 <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">是否启用管理员认证码：</td>
      <td class="forumRowHighlight"><input name="EnableSiteManageCode" type="radio" value="True" <%if EnableSiteManageCode = True then%> checked<%end if%>>
        启用
        <input name="EnableSiteManageCode" type="radio" value="False" <%if EnableSiteManageCode = False then%> checked<%end if%>>
        不启用 <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">管理员认证码：</td>
      <td class="forumRowHighlight"><input name="SiteManageCode" type="text" id="SiteManageCode" style="width: 80" value="<%=SiteManageCode%>" maxlength="6">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">页面防刷新时间：</td>
      <td class="forumRowHighlight"><input name="Refresh" type="text" id="Refresh" style="width: 80" value="<%=Refresh%>">
        秒 <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">新闻模块单页调用条数：</td>
      <td class="forumRowHighlight"><input name="NewInfo" type="text" id="NewInfo" style="width: 80" value="<%=NewInfo%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">产品模块单页调用条数：</td>
      <td class="forumRowHighlight"><input name="ProInfo" type="text" id="ProInfo" style="width: 80" value="<%=ProInfo%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">人才模块单页调用条数：</td>
      <td class="forumRowHighlight"><input name="JobInfo" type="text" id="JobInfo" style="width: 80" value="<%=JobInfo%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">下载模块单页调用条数：</td>
      <td class="forumRowHighlight"><input name="DownInfo" type="text" id="DownInfo" style="width: 80" value="<%=DownInfo%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">其他信息单页调用条数：</td>
      <td class="forumRowHighlight"><input name="OtherInfo" type="text" id="OtherInfo" style="width: 80" value="<%=OtherInfo%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">留言模块单页调用条数：</td>
      <td class="forumRowHighlight"><input name="MessageInfo" type="text" id="MessageInfo" style="width: 80" value="<%=MessageInfo%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">是否生成静态页面：</td>
      <td class="forumRowHighlight"><input name="ISHTML" type="radio" value="1" <%if ISHTML = 1 then%> checked<%end if%>>
        生成静态HTML
        <input name="ISHTML" type="radio" value="0" <%if ISHTML = 0 then%> checked<%end if%>>
        不生成 <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">生成静态页面后缀：</td>
      <td class="forumRowHighlight"><select style="width: 80" name="HTMLName">
          <option value="html" <%if HTMLName="html" then response.write "selected"%>>html</option>
          <option value="htm" <%if HTMLName="htm" then response.write "selected"%>>htm</option>
          <option value="shtml" <%if HTMLName="shtml" then response.write "selected"%>>shtml</option>
          <option value="asp" <%if HTMLName="asp" then response.write "selected"%>>asp</option>
        </select>
        <font color="red">*</font> 如：New.html中“html”即为后缀。可设置如：html、htm、shtml、asp</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">新闻分类生成前缀：</td>
      <td class="forumRowHighlight"><input name="NewSortName" type="text" id="NewSortName" style="width: 180" value="<%=NewSortName%>">
        <font color="red">*</font> 如：New-1.html中“New”即为前缀</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">新闻详细页生成前缀：</td>
      <td class="forumRowHighlight"><input name="NewName" type="text" id="NewName" style="width: 180" value="<%=NewName%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">产品分类生成前缀：</td>
      <td class="forumRowHighlight"><input name="ProSortName" type="text" id="ProSortName" style="width: 180" value="<%=ProSortName%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">产品详细页生成前缀：</td>
      <td class="forumRowHighlight"><input name="ProName" type="text" id="ProName" style="width: 180" value="<%=ProName%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">下载分类生成前缀：</td>
      <td class="forumRowHighlight"><input name="DownSortName" type="text" id="DownSortName" style="width: 180" value="<%=DownSortName%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">下载详细页生成前缀：</td>
      <td class="forumRowHighlight"><input name="DownNameDiy" type="text" id="DownNameDiy" style="width: 180" value="<%=DownNameDiy%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">其他信息分类生成前缀：</td>
      <td class="forumRowHighlight"><input name="OtherSortName" type="text" id="OtherSortName" style="width: 180" value="<%=OtherSortName%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">其他信息详细页生成前缀：</td>
      <td class="forumRowHighlight"><input name="OtherName" type="text" id="OtherName" style="width: 180" value="<%=OtherName%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">人才分类生成前缀：</td>
      <td class="forumRowHighlight"><input name="JobSortName" type="text" id="JobSortName" style="width: 180" value="<%=JobSortName%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">人才详细页生成前缀：</td>
      <td class="forumRowHighlight"><input name="JobNameDiy" type="text" id="JobNameDiy" style="width: 180" value="<%=JobNameDiy%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">企业信息生成前缀：</td>
      <td class="forumRowHighlight"><input name="AboutNameDiy" type="text" id="AboutNameDiy" style="width: 180" value="<%=AboutNameDiy%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">产品询价生成前缀：</td>
      <td class="forumRowHighlight"><input name="AdvisoryNameDiy" type="text" id="AdvisoryNameDiy" style="width: 180" value="<%=AdvisoryNameDiy%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">分隔符：</td>
      <td class="forumRowHighlight"><input name="Separated" type="text" id="Separated" style="width: 80" value="<%=Separated%>">
        <font color="red">*</font> 如：New-1.html中的“-”即为分隔符</td>
    </tr>
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">【浮动客户即时咨询参数】</th>
    </tr>
    <tr>
      <td align="right" class="forumRow">是否启用浮动客户咨询：</td>
      <td class="forumRowHighlight"><input name="JMailDisplay" type="radio" value="1" <%if JMailDisplay = 1 then%> checked<%end if%>>
        启用
        <input name="JMailDisplay" type="radio" value="0" <%if JMailDisplay = 0 then%> checked<%end if%>>
        不启用 <font color="red">*</font> 启用后客户咨询问题会自动记录到后台</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">是否启用同步信箱通知：</td>
      <td class="forumRowHighlight"><input name="JMailPubDisplay" type="radio" value="1" <%if JMailPubDisplay = 1 then%> checked<%end if%>>
        启用
        <input name="JMailPubDisplay" type="radio" value="0" <%if JMailPubDisplay = 0 then%> checked<%end if%>>
        不启用 <font color="red">*</font> 开启此功能后，客户的咨询问题在记录到后台的同时，会同步发送到管理员设置好的信箱。</td>
    </tr>
    <tr>
      <td align="right" class="forumRow">SMTP服务器：</td>
      <td class="forumRowHighlight"><input name="JMailSMTP" type="text" id="JMailSMTP" style="width: 180" value="<%=JMailSMTP%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">SMTP服务器用户名：</td>
      <td class="forumRowHighlight"><input name="JMailUser" type="text" id="JMailUser" style="width: 180" value="<%=JMailUser%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">SMTP服务器密码：</td>
      <td class="forumRowHighlight"><input name="JMailPass" type="text" id="JMailPass" style="width: 180" value="<%=JMailPass%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">发送人：</td>
      <td class="forumRowHighlight"><input name="JMailName" type="text" id="JMailName" style="width: 180" value="<%=JMailName%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">接收信箱：</td>
      <td class="forumRowHighlight"><input name="JMailInFrom" type="text" id="JMailInFrom" style="width: 180" value="<%=JMailInFrom%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">发送信箱：</td>
      <td class="forumRowHighlight"><input name="JMailOutFrom" type="text" id="JMailOutFrom" style="width: 180" value="<%=JMailOutFrom%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow">反馈标题头：</td>
      <td class="forumRowHighlight"><input name="JMailTitle" type="text" id="JMailTitle" style="width: 200" value="<%=JMailTitle%>">
        <font color="red">*</font></td>
    </tr>
    <tr>
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="保存附加参数设置"></td>
    </tr>
  </form>
</table>
<br />
<%
Function SaveSiteInfo()
    If Len(Trim(request.Form("SiteTitle")))<4 Then
        response.Write "<script language='JavaScript'>alert('请详细填写您的网站标题并保持至少在两个汉字以上！');" & "history.back()" & "</script>"
        response.End
    End If
    If Len(Trim(request.Form("SiteUrl")))<9 Then
        response.Write "<script language='JavaScript'>alert('请详细填写您的公司网址！');" & "history.back()" & "</script>"
        response.End
    End If
    If Len(Trim(request.Form("ComName")))<4 Then
        response.Write "<script language='JavaScript'>alert('请详细填写您的公司名称并保持至少在两个汉字以上！');" & "history.back()" & "</script>"
        response.End
    End If
    If Len(Trim(request.Form("Address")))<4 Then
        response.Write "<script language='JavaScript'>alert('请详细填写您的公司地址并保持至少在两个汉字以上！');" & "history.back()" & "</script>"
        response.End
    End If
    If Len(Trim(request.Form("ZipCode")))<6 Then
        response.Write "<script language='JavaScript'>alert('请详细填写邮政编码并保持至少在6个字符以上！');" & "history.back()" & "</script>"
        response.End
    End If
    If Len(Trim(request.Form("Telephone")))<11 Then
        response.Write "<script language='JavaScript'>alert('请详细填写联系电话并保持至少在11个字符以上！');" & "history.back()" & "</script>"
        response.End
    End If
    If Len(Trim(request.Form("Fax")))<11 Then
        response.Write "<script language='JavaScript'>alert('请详细填写传真号码并保持至少在11个字符以上！');" & "history.back()" & "</script>"
        response.End
    End If
    If Len(Trim(request.Form("Email")))<6 Then
        response.Write "<script language='JavaScript'>alert('请详细填写电子邮箱地址并保持至少在6个字符以上！');" & "history.back()" & "</script>"
        response.End
    End If
    If Len(Trim(request.Form("Keywords")))>250 Then
        response.Write "<script language='JavaScript'>alert('请详细填写网站关键词并保持在250个字符以内！');" & "history.back()" & "</script>"
        response.End
    End If
    If Len(Trim(request.Form("Descriptions")))>250 Then
        response.Write "<script language='JavaScript'>alert('请详细填写网站描述并保持在250个字符以内！');" & "history.back()" & "</script>"
        response.End
    End If
    Call SiteInfo()
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select top 1 * from Qianbo_Site"
    rs.Open sql, conn, 1, 3
    rs("SiteTitle") = Trim(Request.Form("SiteTitle"))
    rs("SiteUrl") = Trim(Request.Form("SiteUrl"))
    rs("ComName") = Trim(Request.Form("ComName"))
    rs("Address") = Trim(Request.Form("Address"))
    rs("ZipCode") = Trim(Request.Form("ZipCode"))
    rs("Telephone") = Trim(Request.Form("Telephone"))
    rs("Fax") = Trim(Request.Form("Fax"))
    rs("Email") = Trim(Request.Form("Email"))
    rs("Keywords") = Trim(Request.Form("Keywords"))
    rs("Descriptions") = Trim(Request.Form("Descriptions"))
    rs("IcpNumber") = Trim(Request.Form("IcpNumber"))
    rs("Video") = Trim(Request.Form("Video"))
    rs("SiteLogo") = Trim(Request.Form("SiteLogo"))
    rs("SiteDetail") = Trim(Request.Form("SiteDetail"))
    If Request.Form("MesViewFlag") = 1 Then
        rs("MesViewFlag") = Request.Form("MesViewFlag")
        Conn.Execute "alter table Qianbo_Message alter column ViewFlag bit default 1"
    Else
        rs("MesViewFlag") = 0
        Conn.Execute "alter table Qianbo_Message alter column ViewFlag bit default 0"
    End If
    rs.update
    rs.Close
    Set rs = Nothing
    response.Write "<script language='javascript'>alert('系统主参数设置成功！');location.replace('SetSite.Asp');</script>"
End Function

Function ViewSiteInfo()
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select top 1 * from Qianbo_Site"
    rs.Open sql, conn, 1, 1
    If rs.bof And rs.EOF Then
        response.Write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>数据库记录读取错误！</font></div>")
        response.End
    Else
        SiteTitle = rs("SiteTitle")
        SiteUrl = rs("SiteUrl")
        ComName = rs("ComName")
        Address = rs("Address")
        ZipCode = rs("ZipCode")
        Telephone = rs("Telephone")
        Fax = rs("Fax")
        Email = rs("Email")
        Keywords = rs("Keywords")
        Descriptions = rs("Descriptions")
        IcpNumber = rs("IcpNumber")
        MesViewFlag = rs("MesViewFlag")
        Video = rs("Video")
        SiteLogo = rs("SiteLogo")
        SiteDetail = rs("SiteDetail")
        rs.Close
        Set rs = Nothing
    End If
End Function

Function SaveConstInfo()
    Call SiteInfo()
    Set ASM = server.CreateObject("Adod" & "b.St" & "ream")
    ASM.Type = 2
    ASM.mode = 3
    ASM.charset = "utf-8"
    ASM.Open
    ASM.WriteText "<!--#include file=""Charset.Asp"" -->" & vbCrLf
    ASM.WriteText "<" & "%" & vbCrLf
    ASM.WriteText "Const SysRootDir = " & Chr(34) & ReplaceConstChar(Trim(request("SysRootDir"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const SiteDataPath = " & Chr(34) & ReplaceConstChar(Trim(request("SiteDataPath"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const EnableStopInjection = " & ReplaceConstChar(Trim(request("EnableStopInjection"))) & "" & vbCrLf
    ASM.WriteText "Const EnableSiteManageCode = " & ReplaceConstChar(Trim(request("EnableSiteManageCode"))) & "" & vbCrLf
    ASM.WriteText "Const SiteManageCode = " & Chr(34) & ReplaceConstChar(Trim(request("SiteManageCode"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const Refresh = " & ReplaceConstChar(Trim(request("Refresh"))) & "" & vbCrLf
    ASM.WriteText "Const NewInfo = " & ReplaceConstChar(Trim(request("NewInfo"))) & "" & vbCrLf
    ASM.WriteText "Const ProInfo = " & ReplaceConstChar(Trim(request("ProInfo"))) & "" & vbCrLf
    ASM.WriteText "Const JobInfo = " & ReplaceConstChar(Trim(request("JobInfo"))) & "" & vbCrLf
    ASM.WriteText "Const DownInfo = " & ReplaceConstChar(Trim(request("DownInfo"))) & "" & vbCrLf
    ASM.WriteText "Const OtherInfo = " & ReplaceConstChar(Trim(request("OtherInfo"))) & "" & vbCrLf
    ASM.WriteText "Const MessageInfo = " & ReplaceConstChar(Trim(request("MessageInfo"))) & "" & vbCrLf
    ASM.WriteText "Const ISHTML = " & ReplaceConstChar(Trim(request("ISHTML"))) & "" & vbCrLf
    ASM.WriteText "Const HTMLName = " & Chr(34) & ReplaceConstChar(Trim(request("HTMLName"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const NewSortName = " & Chr(34) & ReplaceConstChar(Trim(request("NewSortName"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const NewName = " & Chr(34) & ReplaceConstChar(Trim(request("NewName"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const ProSortName = " & Chr(34) & ReplaceConstChar(Trim(request("ProSortName"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const ProName = " & Chr(34) & ReplaceConstChar(Trim(request("ProName"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const DownSortName = " & Chr(34) & ReplaceConstChar(Trim(request("DownSortName"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const DownNameDiy = " & Chr(34) & ReplaceConstChar(Trim(request("DownNameDiy"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const OtherSortName = " & Chr(34) & ReplaceConstChar(Trim(request("OtherSortName"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const OtherName = " & Chr(34) & ReplaceConstChar(Trim(request("OtherName"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const JobSortName = " & Chr(34) & ReplaceConstChar(Trim(request("JobSortName"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const JobNameDiy = " & Chr(34) & ReplaceConstChar(Trim(request("JobNameDiy"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const AboutNameDiy = " & Chr(34) & ReplaceConstChar(Trim(request("AboutNameDiy"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const AdvisoryNameDiy = " & Chr(34) & ReplaceConstChar(Trim(request("AdvisoryNameDiy"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const Separated = " & Chr(34) & ReplaceConstChar(Trim(request("Separated"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const JMailDisplay = " & ReplaceConstChar(Trim(request("JMailDisplay"))) & "" & vbCrLf
    ASM.WriteText "Const JMailPubDisplay = " & ReplaceConstChar(Trim(request("JMailPubDisplay"))) & "" & vbCrLf
    ASM.WriteText "Const JMailSMTP = " & Chr(34) & ReplaceConstChar(Trim(request("JMailSMTP"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const JMailUser = " & Chr(34) & ReplaceConstChar(Trim(request("JMailUser"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const JMailPass = " & Chr(34) & ReplaceConstChar(Trim(request("JMailPass"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const JMailName = " & Chr(34) & ReplaceConstChar(Trim(request("JMailName"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const JMailInFrom = " & Chr(34) & ReplaceConstChar(Trim(request("JMailInFrom"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const JMailOutFrom = " & Chr(34) & ReplaceConstChar(Trim(request("JMailOutFrom"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "Const JMailTitle = " & Chr(34) & ReplaceConstChar(Trim(request("JMailTitle"))) & Chr(34) & "" & vbCrLf
    ASM.WriteText "%" & ">"
    ASM.SaveToFile Server.mappath("../Include/Const.Asp"), 2
    ASM.flush
    ASM.Close
    Set ASM = Nothing
    If ReplaceConstChar(Trim(request("ISHTML"))) = 0 Then Call FsoDelHtml(ReplaceConstChar(Trim(request("HTMLName"))))
    response.Write "<script language=""javascript"">alert('系统附加参数设置成功！');location.href='SetSite.Asp';</script>"
End Function

Sub FsoDelHtml(HTMLName)
    Dim Fso, FsoOut, File
    Set Fso = Server.CreateObject("Scripting.FileSystemObject")
    Set FsoOut = Fso.GetFolder(Server.Mappath(SysRootDir))
    For Each File In FsoOut.Files
        If Right(File.Name, Len(File.Name) - InStrRev(File.Name, ".")) = HTMLName And HTMLName <> "asp" Then
            Response.Write "<span style=""color:red; padding-left: 18px"">" & File.Name & "</span> 成功删除…<br />"
            Fso.DeleteFile File.Path, True
        End If
    Next
    Set FsoOut = Nothing
    Set Fso = Nothing
End Sub
%>