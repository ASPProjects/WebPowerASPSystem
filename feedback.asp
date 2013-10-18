<!--#include file="./____Core.asp" -->
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
<!--#include file="template/jielansi/page/feedback.asp"-->