<!--#include file="CheckAdmin.Asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="AdminDefaultTheme/index.css" type="text/css" rel="stylesheet" />
<link href="AdminDefaultTheme/MasterPage.css" type="text/css" rel="stylesheet" />
<link href="AdminDefaultTheme/Guide.css" type="text/css" rel="stylesheet" />
<title>管理导航菜单</title>
<script type="text/javascript">
<!--
function Creathtml()
{
    var bln=confirm("注意：添加、修改、删除相关数据时会自动生成、更新、删除所生成的静态文件。\n如果您没有对模板作过修改，不需要批量生成所有商品或新闻详细页面！\n如果您仅对产品、新闻、下载、人才等分类页面作过修改，只需要生成相关分类页面。\n\n请确定是否操作？");
    return bln;
}
function CreathtmlAll()
{
    var bln=confirm("警告：批量生成全站静态页面将耗费较多系统资源！\n\n请确定是否操作？");
    return bln;
}
-->
</script>
</head>

<body id="Guidebody">
<div id="Guide_back">
  <ul>
    <li id="Guide_top">
      <div id="Guide_toptext">快捷导航</div>
    </li>
    <li id="Guide_main">
      <div id="Guide_box">
        <div class="guide">
          <ul id="Links">
            <%
Select Case Request("ID")
    Case "System"
        Call System()
    Case "News"
        Call News()
    Case "Product"
        Call Product()
    Case "DownLoad"
        Call DownLoad()
    Case "Talent"
        Call Talent()
    Case "Other"
        Call Other()
    Case "Feedback"
        Call Feedback()
    Case "User"
        Call User()
    Case "Html"
        Call Html()
    Case Else
        Call System()
End Select
%>
            <%Sub System%>
            <div class="guideexpand">网站参数设置</div>
            <div class="guide">
              <ul id="Links">
                <li><a href="SetSite.Asp" target="main_right">网站参数设置</a></li>
                <li><a href="NavigationEdit.Asp?Result=Add" target="main_right">导航栏添加</a></li>
                <li><a href="NavigationList.Asp" target="main_right">导航栏管理</a></li>
                <li><a href="FriendLinkEdit.Asp?Result=Add" target="main_right">友情链接添加</a></li>
                <li><a href="FriendLinkList.Asp" target="main_right">友情链接管理</a></li>
                <li><a href="SetKey.Asp" target="main_right">站内链接管理</a></li>
                <li><a href="LinkEdit.Asp?Result=Add" target="main_right">站内链接添加</a></li>
                <li><a href="SQL.Asp" target="main_right">数据库结构管理</a></li>
                <li><a href="Admin_SiteMap.Asp" target="main_right">生成谷歌SiteMap</a></li>
                <li><a href="UserMessage.Asp" target="main_right">客户即时咨询管理</a></li>
              </ul>
            </div>
            <div class="guideexpand">网站推广入口</div>
            <div class="guide">
              <ul id="Links">
                <li><a href="http://www.baidu.com/search/url_submit.html" target="_blank">百度登录入口</a></li>
                <li><a href="http://www.google.com/intl/zh-CN/add_url.html" target="_blank">Google登录入口</a></li>
                <li><a href="http://search.help.cn.yahoo.com/h4_4.html" target="_blank">Yahoo登录入口</a></li>
                <li><a href="http://search.msn.com/docs/submit.Aspx" target="_blank">Live登录入口</a></li>
                <li><a href="http://www.dmoz.org/World/Chinese_Simplified/" target="_blank">Dmoz登录入口</a></li>
                <li><a href="http://www.alexa.com/site/help/webmasters" target="_blank">Alexa登录入口</a></li>
                <li><a href="http://ads.zhongsou.com/register/page.jsp" target="_blank">中搜登录入口</a></li>
                <li><a href="http://iask.com/guest/add_url.php" target="_blank">爱问登录入口</a></li>
              </ul>
            </div>
            <%End Sub%>
            <%Sub News%>
            <li><a href="NewsSort.Asp?Action=Add&ParentID=0" target="main_right">新闻类别管理</a></li>
            <li><a href="NewsList.Asp" target="main_right">新闻列表管理</a></li>
            <li><a href="NewsEdit.Asp?Result=Add" target="main_right">添加新闻</a></li>
            <li><a href="AboutList.Asp" target="main_right">企业信息列表</a></li>
            <li><a href="AboutEdit.Asp?Result=Add" target="main_right">添加企业信息</a></li>
            <%End Sub%>
            <%Sub Product%>
            <li><a href="ProductSort.Asp?Action=Add&ParentID=0" target="main_right">产品类别管理</a></li>
            <li><a href="ProductList.Asp" target="main_right">产品列表管理</a></li>
            <li><a href="ProductEdit.Asp?Result=Add" target="main_right">添加产品信息</a></li>
            <%End Sub%>
            <%Sub DownLoad%>
            <li><a href="DownSort.Asp?Action=Add&ParentID=0" target="main_right">下载类别管理</a></li>
            <li><a href="DownList.Asp" target="main_right">下载列表管理</a></li>
            <li><a href="DownEdit.Asp?Result=Add" target="main_right">添加下载信息</a></li>
            <%End Sub%>
            <%Sub Talent%>
            <li><a href="JobsList.Asp" target="main_right">招聘列表管理</a></li>
            <li><a href="JobsEdit.Asp?Result=Add" target="main_right">添加招聘信息</a></li>
            <%End Sub%>
            <%Sub Other%>
            <li><a href="OthersSort.Asp?Action=Add&ParentID=0" target="main_right">信息类别管理</a></li>
            <li><a href="OthersList.Asp" target="main_right">信息列表管理</a></li>
            <li><a href="OthersEdit.Asp?Result=Add" target="main_right">添加信息</a></li>
            <%End Sub%>
            <%Sub Feedback%>
            <li><a href="MessageList.Asp" target="main_right">留言信息管理</a></li>
            <li><a href="OrderList.Asp" target="main_right">订单信息管理</a></li>
            <li><a href="TalentsList.Asp" target="main_right">人才信息管理</a></li>
            <%End Sub%>
            <%Sub User%>
            <li><a href="AdminList.Asp" target="main_right">网站管理员管理</a></li>
            <li><a href="AdminEdit.Asp?Result=Add" target="main_right">添加网站管理员</a></li>
            <li><a href="MemList.Asp" target="main_right">前台会员资料</a></li>
            <li><a href="MemGroup.Asp" target="main_right">会员组别管理</a></li>
            <li><a href="MemGroup.Asp?Result=Add" target="main_right">添加会员组别</a></li>
            <li><a href="ManageLog.Asp" target="main_right">后台登录日志管理</a></li>
            <%End Sub%>
            <%Sub Html%>
            <li><a href="Admin_htmlsort.Asp" target="main_right">生成产品分类页面</a></li>
            <li><a href="Admin_htmlpro.Asp" target="main_right" onClick="return Creathtml()">生成所有产品详细页面</a></li>
            <li><a href="Admin_htmlnewsort.Asp" target="main_right">生成新闻分类页面</a></li>
            <li><a href="Admin_htmlnews.Asp" target="main_right" onClick="return Creathtml()">生成新闻详细页面</a></li>
            <li><a href="Admin_htmlinfo.Asp" target="main_right">生成企业信息列表</a></li>
            <li><a href="Admin_htmldownsort.Asp" target="main_right">生成下载分类页面</a></li>
            <li><a href="Admin_htmldown.Asp" target="main_right" onClick="return Creathtml()">生成下载详细页面</a></li>
            <li><a href="Admin_htmljobsort.Asp" target="main_right">生成人才招聘列表</a></li>
            <li><a href="Admin_htmljob.Asp" target="main_right">生成人才招聘详细页面</a></li>
            <li><a href="Admin_htmlothersort.Asp" target="main_right">生成其他信息分类页面</a></li>
            <li><a href="Admin_htmlother.Asp" target="main_right" onClick="return Creathtml()">生成其他信息详细页面</a></li>
            <li><a href="Admin_htmlindex.Asp" target="main_right">生成首页|询价|其他</a></li>
            <li><a href="Admin_html.Asp" target="main_right" onClick="return CreathtmlAll()"><font color="red">生成全站静态页面</font></a></li>
            <%End Sub%>
          </ul>
        </div>
      </div>
    </li>
    <li id="Guide_bottom"></li>
  </ul>
</div>
</body>
</html>