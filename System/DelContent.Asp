<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
Dim Result, Selectid
Dim ix, tempx, selectx, len_idx
Result = request.QueryString("Result")
SelectID = request.Form("SelectID")
Call SiteInfo()
Select Case Result
    Case "SiteLink"
        Select Case Trim(Request.Form("batch"))
        Case "批量生效"
            If SelectID<>"" Then conn.Execute "update Qianbo_Link set State = 1 where ID in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "批量失效"
            If SelectID<>"" Then conn.Execute "update Qianbo_Link set State = 0 where ID in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "删除所选"
            If SelectID<>"" Then conn.Execute "delete from Qianbo_Link where ID in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
    End Select
    Case "Administrators"
        Select Case Trim(Request.Form("batch"))
        Case "批量生效"
            If SelectID<>"" Then conn.Execute "update Qianbo_Admin set Working = 1 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "批量失效"
            If SelectID<>"" Then conn.Execute "update Qianbo_Admin set Working = 0 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "删除所选"
            If SelectID<>"" Then conn.Execute "delete from Qianbo_Admin where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
    End Select
    Case "LoginLog"
        If SelectID<>"" Then conn.Execute "delete from Qianbo_AdminLog where id in ("&SelectID&")"
        response.redirect request.servervariables("http_referer")
    Case "MemGroup"
        If SelectID<>"" Then conn.Execute "delete from Qianbo_MemGroup where GroupID in ('"&SelectID&"')"
        conn.Execute "Update Qianbo_About set GroupID='200603281858588888' , Exclusive='>=' where GroupID = '"&SelectID&"'"
        conn.Execute "Update Qianbo_Download set GroupID='200603281858588888' , Exclusive='>=' where GroupID = '"&SelectID&"'"
        conn.Execute "Update Qianbo_Members set GroupID='200603281858588888' , GroupName='临时游客' where GroupID = '"&SelectID&"'"
        conn.Execute "Update Qianbo_News set GroupID='200603281858588888' , Exclusive='>=' where GroupID = '"&SelectID&"'"
        conn.Execute "Update Qianbo_Others set GroupID='200603281858588888' , Exclusive='>=' where GroupID = '"&SelectID&"'"
        conn.Execute "Update Qianbo_Products set GroupID='200603281858588888' , Exclusive='>=' where GroupID = '"&SelectID&"'"
        response.redirect request.servervariables("http_referer")
    Case "Members"
        Select Case Trim(Request.Form("batch"))
        Case "批量生效"
            If SelectID<>"" Then conn.Execute "update Qianbo_Members set Working = 1 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "批量失效"
            If SelectID<>"" Then conn.Execute "update Qianbo_Members set Working = 0 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "删除所选"
            If SelectID<>"" Then conn.Execute "delete from Qianbo_Members where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
    End Select
    Case "About"
        If SelectID<>"" Then conn.Execute "delete from Qianbo_About where id in ("&SelectID&")"
        response.redirect request.servervariables("http_referer")
    Case "Products"
        Select Case Trim(Request.Form("batch"))
        Case "批量修改产品参数"
            pro_id = request("pro_id")
            ProductNo = request("ProductNo")
            ProductName = request("ProductName")
            N_Price = request("N_Price")
            P_Price = request("P_Price")
            ProductModel = request("ProductModel")
            Stock = request("Stock")
            arr_id = Split(Replace(pro_id, " ", ""), ",")
            arr_ProductNo = Split(Replace(ProductNo, " ", ""), ",")
            arr_ProductName = Split(Replace(ProductName, " ", ""), ",")
            arr_N_Price = Split(Replace(N_Price, " ", ""), ",")
            arr_P_Price = Split(Replace(P_Price, " ", ""), ",")
            arr_ProductModel = Split(Replace(ProductModel, " ", ""), ",")
            arr_Stock = Split(Replace(Stock, " ", ""), ",")
            For i = 0 To UBound(arr_id)
                conn.Execute("update Qianbo_Products set ProductNo='"&arr_ProductNo(i)&"' where id="&arr_id(i)&"")
                conn.Execute("update Qianbo_Products set ProductName='"&arr_ProductName(i)&"' where id="&arr_id(i)&"")
                conn.Execute("update Qianbo_Products set N_Price="&arr_N_Price(i)&" where id="&arr_id(i)&"")
                conn.Execute("update Qianbo_Products set P_Price="&arr_P_Price(i)&" where id="&arr_id(i)&"")
                conn.Execute("update Qianbo_Products set ProductModel='"&arr_ProductModel(i)&"' where id="&arr_id(i)&"")
                conn.Execute("update Qianbo_Products set Stock="&arr_Stock(i)&" where id="&arr_id(i)&"")
            Next
            response.redirect request.servervariables("http_referer")
        Case "批量生效"
            If SelectID<>"" Then conn.Execute "update Qianbo_Products set ViewFlag = 1 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "批量失效"
            If SelectID<>"" Then conn.Execute "update Qianbo_Products set ViewFlag = 0 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "删除所选"
            If SelectID<>"" Then conn.Execute "delete from Qianbo_Products where id in ("&SelectID&")"
            If ISHTML = 1 Then
                ix = 0
                SelectID = Replace(SelectID, " ", "")
                len_idx = Len(SelectID)
                tempx = tempx&Left(SelectID, 1)
                Do While ix<len_idx
                    ix = ix + 1
                    tempx = Left(SelectID, 1)
                    If tempx = "," Then
                        Call DoDelslhtml(""&ProName&""&Separated&""&selectx&"."&HTMLName&"")
                        selectx = ""
                    Else
                        selectx = selectx&tempx
                    End If
                    If len_idx - ix = 0 Then Call DoDelslhtml(""&ProName&""&Separated&""&selectx&"."&HTMLName&"")
                    SelectID = Right(SelectID, len_idx - ix)
                Loop
            End If
            response.redirect request.servervariables("http_referer")
    End Select
    Case "News"
        Select Case Trim(Request.Form("batch"))
        Case "批量生效"
            If SelectID<>"" Then conn.Execute "update Qianbo_News set ViewFlag = 1 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "批量失效"
            If SelectID<>"" Then conn.Execute "update Qianbo_News set ViewFlag = 0 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "删除所选"
            If SelectID<>"" Then conn.Execute "delete from Qianbo_News where id in ("&SelectID&")"
            If ISHTML = 1 Then
                ix = 0
                SelectID = Replace(SelectID, " ", "")
                len_idx = Len(SelectID)
                tempx = tempx&Left(SelectID, 1)
                Do While ix<len_idx
                    ix = ix + 1
                    tempx = Left(SelectID, 1)
                    If tempx = "," Then
                        Call DoDelslhtml(""&NewName&""&Separated&""&selectx&"."&HTMLName&"")
                        selectx = ""
                    Else
                        selectx = selectx&tempx
                    End If
                    If len_idx - ix = 0 Then Call DoDelslhtml(""&NewName&""&Separated&""&selectx&"."&HTMLName&"")
                    SelectID = Right(SelectID, len_idx - ix)
                Loop
            End If
            response.redirect request.servervariables("http_referer")
    End Select
    Case "UserMessage"
        Select Case Trim(Request.Form("batch"))
        Case "批量处理"
            If SelectID<>"" Then conn.Execute "update Qianbo_Biz set BizOK = 1 where ID in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "删除所选"
            If SelectID<>"" Then conn.Execute "delete from Qianbo_Biz where ID in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
    End Select
    Case "Download"
        Select Case Trim(Request.Form("batch"))
        Case "批量生效"
            If SelectID<>"" Then conn.Execute "update Qianbo_Download set ViewFlag = 1 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "批量失效"
            If SelectID<>"" Then conn.Execute "update Qianbo_Download set ViewFlag = 0 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "删除所选"
            If SelectID<>"" Then conn.Execute "delete from Qianbo_Download where id in ("&SelectID&")"
            If ISHTML = 1 Then
                ix = 0
                SelectID = Replace(SelectID, " ", "")
                len_idx = Len(SelectID)
                tempx = tempx&Left(SelectID, 1)
                Do While ix<len_idx
                    ix = ix + 1
                    tempx = Left(SelectID, 1)
                    If tempx = "," Then
                        Call DoDelslhtml(""&DownNameDiy&""&Separated&""&selectx&"."&HTMLName&"")
                        selectx = ""
                    Else
                        selectx = selectx&tempx
                    End If
                    If len_idx - ix = 0 Then Call DoDelslhtml(""&DownNameDiy&""&Separated&""&selectx&"."&HTMLName&"")
                    SelectID = Right(SelectID, len_idx - ix)
                Loop
            End If
            response.redirect request.servervariables("http_referer")
    End Select
    Case "Jobs"
        Select Case Trim(Request.Form("batch"))
        Case "批量生效"
            If SelectID<>"" Then conn.Execute "update Qianbo_Jobs set ViewFlag = 1 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "批量失效"
            If SelectID<>"" Then conn.Execute "update Qianbo_Jobs set ViewFlag = 0 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "删除所选"
            If SelectID<>"" Then conn.Execute "delete from Qianbo_Jobs where id in ("&SelectID&")"
            If ISHTML = 1 Then
                ix = 0
                SelectID = Replace(SelectID, " ", "")
                len_idx = Len(SelectID)
                tempx = tempx&Left(SelectID, 1)
                Do While ix<len_idx
                    ix = ix + 1
                    tempx = Left(SelectID, 1)
                    If tempx = "," Then
                        Call DoDelslhtml(""&JobNameDiy&""&Separated&""&selectx&"."&HTMLName&"")
                        selectx = ""
                    Else
                        selectx = selectx&tempx
                    End If
                    If len_idx - ix = 0 Then Call DoDelslhtml(""&JobNameDiy&""&Separated&""&selectx&"."&HTMLName&"")
                    SelectID = Right(SelectID, len_idx - ix)
                Loop
            End If
            response.redirect request.servervariables("http_referer")
    End Select
    Case "Message"
        Select Case Trim(Request.Form("batch"))
        Case "批量生效"
            If SelectID<>"" Then conn.Execute "update Qianbo_Message set ViewFlag = 1 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "批量失效"
            If SelectID<>"" Then conn.Execute "update Qianbo_Message set ViewFlag = 0 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "删除所选"
            If SelectID<>"" Then conn.Execute "delete from Qianbo_Message where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
    End Select
    Case "Order"
        If SelectID<>"" Then conn.Execute "delete from Qianbo_Order where id in ("&SelectID&")"
        response.redirect request.servervariables("http_referer")
    Case "Talents"
        If SelectID<>"" Then conn.Execute "delete from Qianbo_Talents where id in ("&SelectID&")"
        response.redirect request.servervariables("http_referer")
    Case "Navigation"
        Select Case Trim(Request.Form("batch"))
        Case "批量生效"
            If SelectID<>"" Then conn.Execute "update Qianbo_Navigation set ViewFlag = 1 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "批量失效"
            If SelectID<>"" Then conn.Execute "update Qianbo_Navigation set ViewFlag = 0 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "删除所选"
            If SelectID<>"" Then conn.Execute "delete from Qianbo_Navigation where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
    End Select
    Case "FriendLink"
        Select Case Trim(Request.Form("batch"))
        Case "批量生效"
            If SelectID<>"" Then conn.Execute "update Qianbo_FriendLink set ViewFlag = 1 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "批量失效"
            If SelectID<>"" Then conn.Execute "update Qianbo_FriendLink set ViewFlag = 0 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "删除所选"
            If SelectID<>"" Then conn.Execute "delete from Qianbo_FriendLink where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
    End Select
    Case "Others"
        Select Case Trim(Request.Form("batch"))
        Case "批量生效"
            If SelectID<>"" Then conn.Execute "update Qianbo_Others set ViewFlag = 1 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "批量失效"
            If SelectID<>"" Then conn.Execute "update Qianbo_Others set ViewFlag = 0 where id in ("&SelectID&")"
            response.redirect request.servervariables("http_referer")
        Case "删除所选"
            If SelectID<>"" Then conn.Execute "delete from Qianbo_Others where id in ("&SelectID&")"
            If ISHTML = 1 Then
                ix = 0
                SelectID = Replace(SelectID, " ", "")
                len_idx = Len(SelectID)
                tempx = tempx&Left(SelectID, 1)
                Do While ix<len_idx
                    ix = ix + 1
                    tempx = Left(SelectID, 1)
                    If tempx = "," Then
                        Call DoDelslhtml(""&OtherName&""&Separated&""&selectx&"."&HTMLName&"")
                        selectx = ""
                    Else
                        selectx = selectx&tempx
                    End If
                    If len_idx - ix = 0 Then Call DoDelslhtml(""&OtherName&""&Separated&""&selectx&"."&HTMLName&"")
                    SelectID = Right(SelectID, len_idx - ix)
                Loop
            End If
            response.redirect request.servervariables("http_referer")
    End Select
    Case Else
End Select
%>