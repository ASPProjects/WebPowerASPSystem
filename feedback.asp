<!--#include file="Include/Const.Asp" -->
<!--#include file="Include/NoSQL.Asp" -->
<!--#include file="Include/ConnSiteData.Asp" -->
<!--#include file="dy.Asp" -->
<%
Call SiteInfo()
Dim MesName, Content, SecretFlag, mMemID, mLinkman, mSex, mCompany, mAddress, mZipCode, mTelephone, mFax, mMobile, mEmail
If session("MemName")<>"" And session("MemLogin") = "Succeed" Then
    Call MemInfo()
Else
    mSex = "先生"
    mMemID = 0
    SecretFlag = "<font color=""red"">会员功能，请先注册会员。</font>"
End If
%>
<!--#include file="template/default/common/header_common.asp"-->
<div id="ct" class="wp cl">
<!--#include file="template/default/common/sider.asp"-->
<div id="mn">
<span class="tit">信息反馈</span>
<div id="aboutus">

<!--#include file="ly.asp"-->

</div>
</div></div>
<!--#include file="template/default/common/footer_common.asp"-->