<%
CompanyAdminName = ReplaceBadChar(Trim(Request.Cookies("CompanyZY")("AdminName")))
CompanyAdminPurview = ReplaceBadChar(Trim(Request.Cookies("CompanyZY")("AdminPurview")))
CompanyLoginSystem = ReplaceBadChar(Trim(Request.Cookies("CompanyZY")("LoginSystem")))
CompanyAdminLoginCode = ReplaceBadChar(Trim(Request.Cookies("CompanyZY")("AdminLoginCode")))
CompanyAdminLoginIP = ReplaceBadChar(Trim(Request.Cookies("CompanyZY")("LoginIP")))

Function ReplaceBadChar(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceBadChar = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "+,',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ",--"
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    tempChar = Replace(tempChar, "@@", "@")
    ReplaceBadChar = tempChar
End Function

Dim AdminAction
AdminAction = request.QueryString("AdminAction")
Select Case AdminAction
    Case "Out"
        Call OutLogin()
    Case Else
        Call Login()
End Select

Sub Login()
    If CompanyAdminName = "" Or CompanyAdminPurview = "" Or CompanyLoginSystem <> "Succeed" Or (EnableSiteManageCode = True And CompanyAdminLoginCode <> SiteManageCode) Then
        Response.Redirect "Admin_Login.Asp"
        Response.End
    End If
End Sub

Sub OutLogin()
    Response.Cookies("CompanyZY")("AdminName") = ""
    Response.Cookies("CompanyZY")("AdminPurview") = ""
    Response.Cookies("CompanyZY")("LoginSystem") = ""
    Response.Cookies("CompanyZY")("AdminLoginCode") = ""
    Response.Cookies("CompanyZY")("LoginIP") = ""
    Response.Redirect("Admin_Login.Asp")
End Sub
%>