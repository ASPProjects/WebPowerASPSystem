<!--#include file="./____Core.asp" -->
<%
If DateDiff("s", session("time"), Now())<Refresh Then
    response.Write "<script language='JavaScript'>alert('防刷新机制启动：请不要在 "&Refresh&" 秒内重复刷新！');" & "history.back()" & "</script>"
    response.End
Else
    session("time") = Now()
End If
Dim rs, sql
Dim MesName, Content, SecretFlag, mMemID, mLinkman, mSex, mCompany, mAddress, mZipCode, mTelephone, mFax, mMobile, mEmail, CheckCode
MesName = Trim(request.Form("MesName"))
Content = Trim(request.Form("Content"))
If Trim(request.Form("SecretFlag")) = "1" Then
    SecretFlag = 1
Else
    SecretFlag = 0
End If
mMemID = request.QueryString("MemberID")
mLinkman = Trim(request.Form("Linkman"))
mSex = Trim(request.Form("Sex"))
mCompany = Trim(request.Form("Company"))
mAddress = Trim(request.Form("Address"))
mZipCode = Trim(request.Form("ZipCode"))
mTelephone = Trim(request.Form("Telephone"))
mFax = Trim(request.Form("Fax"))
mMobile = Trim(request.Form("Mobile"))
mEmail = Trim(request.Form("Email"))
CheckCode = Trim(request.Form("CheckCode"))
Dim ErrMessage, ErrMsg(9), FindErr(9), i
ErrMsg(0) = "请填写留言主题。"
ErrMsg(1) = "内容不能少于5个字符。"
ErrMsg(2) = "请填写您的称呼。"
ErrMsg(3) = "请正确填写单位名称、联系地址。"
ErrMsg(4) = "请正确填写邮政编码。"
ErrMsg(5) = "请正确填写联系电话、传真号码、手机号码。"
ErrMsg(6) = "电子信箱格式错误。"
ErrMsg(7) = "验证码不能为空，请返回检查。"
ErrMsg(8) = "您在【留言反馈】页面停留的时间过长，导致验证码失效。\n请返回并刷新【留言反馈】页面！"
ErrMsg(9) = "您输入的验证码和系统产生的不一致，请重新输入。"
If Len(MesName)>100 Or Len(MesName) = 0 Then
    FindErr(0) = True
End If
If Len(Content)<5 Then
    FindErr(1) = True
End If

If Len(mCompany)>100 Or Len(mAddress)>100 Or Len(mCompany) = 0 Or Len(mAddress) = 0 Then
    FindErr(3) = True
End If
If Len(mZipCode)<>6 Then
    FindErr(4) = True
End If
If Len(mFax)>50 Or Len(mMobile)>50 Or Len(mFax) = 0 Or Len(mMobile) = 0 Then
    FindErr(5) = True
End If

If CheckCode = "" Then
    FindErr(7) = True
End If
If Trim(Session("CheckCode")) = "" Then
    FindErr(8) = True
End If
If CheckCode <> Session("CheckCode") Then
    FindErr(9) = True
End If
For i = 0 To UBound(FindErr)
    If FindErr(i) = True Then
        ErrMessage = ErrMessage+ErrMsg(i) + "\n"
    End If
Next
If Not (ErrMessage = "" Or IsNull(ErrMessage)) Then
    WriteMsg(ErrMessage)
    response.End
End If
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from Qianbo_Message"
rs.Open sql, conn, 1, 3
rs.addnew
rs("MesName") = StrReplace(MesName)
rs("Content") = StrReplace(Content)
rs("MemID") = mMemID
rs("Linkman") = StrReplace(mLinkman)
rs("Sex") = mSex
rs("Company") = StrReplace(mCompany)
rs("Address") = StrReplace(mAddress)
rs("ZipCode") = StrReplace(mZipCode)
rs("Telephone") = StrReplace(mTelephone)
rs("Fax") = StrReplace(mFax)
rs("Mobile") = StrReplace(mMobile)
rs("Email") = mEmail
rs("SecretFlag") = SecretFlag
rs("AddTime") = Now()
rs.update
rs.Close
Set rs = Nothing
Call SiteInfo()
If MesViewFlag = 0 Then
    response.Write "<script language='javascript'>alert('成功提交，但系统已被设置为审核可见，请等待管理员审核、回复！');location.replace('index.Asp');</script>"
Else
    response.Write "<script language='javascript'>alert('成功提交，请等待管理员回复！');location.replace('index.Asp');</script>"
End If
%>