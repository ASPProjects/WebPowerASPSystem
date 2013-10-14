<%
						
						Function News(SortPath)
    Dim rs, sql, NewsName, NewFlag,AddTime
    Set rs = server.CreateObject("adodb.recordset")
        sql = "select top 6 * from Qianbo_News where ViewFlag and Instr(SortPath,'"&SortPath&"')>0 order by id desc"
    rs.Open sql, conn, 1, 1
  
    If rs.EOF Then
        Response.Write "<center>暂无相关信息</center>"&vbCrLf
    Else
        Do While Not rs.EOF
            If Now() - rs("AddTime")<= 2 Then
                NewsName = StrLeft(rs("NewsName"), 30)
				AddTime = StrLeft(rs("AddTime"), 8)
                NewFlag = " <img src=""Images/new.gif"" align=""absmiddle"">"
            Else
                NewsName = Mid(rs("NewsName"), 1, 10)
                NewFlag = ""
            End If
            If ISHTML = 1 Then
                AutoLink = ""&NewName&""&Separated&""&rs("ID")&"."&HTMLName&""
            Else
                AutoLink = "NewsView.Asp?ID="&rs("id")&""
            End If
			%>
<li><a href="<% =AutoLink%>" title="<% =NewsName%>">
  <% =NewsName%>
  </A><span>[<% =left(rs("addtime"),10)%>]</span></li>
<%
           
            rs.movenext
        Loop
    End If

    rs.Close
    Set rs = Nothing
End Function

						Function News2(SortPath)
    Dim rs, sql, NewsName, NewFlag,AddTime
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select top 9 * from Qianbo_News where ViewFlag and Instr(SortPath,'"&SortPath&"')>0 order by id desc"
    rs.Open sql, conn, 1, 1
  
    If rs.EOF Then
        Response.Write "<center>暂无相关信息</center>"&vbCrLf
    Else
        Do While Not rs.EOF
            If Now() - rs("AddTime")<= 2 Then
                NewsName = StrLeft(rs("NewsName"), 30)
				AddTime = StrLeft(rs("AddTime"), 8)
                NewFlag = " <img src=""Images/new.gif"" align=""absmiddle"">"
            Else
                NewsName = Mid(rs("NewsName"), 1, 12)
                NewFlag = ""
            End If
            If ISHTML = 1 Then
                AutoLink = ""&NewName&""&Separated&""&rs("ID")&"."&HTMLName&""
            Else
                AutoLink = "NewsView.Asp?ID="&rs("id")&""
            End If
			%>
<li><a href="<% =AutoLink%>" title="<% =NewsName%>">
  <% =NewsName%>
  </A></li>
<%
           
            rs.movenext
        Loop
    End If

    rs.Close
    Set rs = Nothing
End Function

Function Products(SortPath, trs, tds)
    Dim rs, sql, tr, td, ProductName, SmallPicPath
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select top "&trs * tds&" * from Qianbo_Products where ViewFlag and CommendFlag order by id desc"
    'sql="select top "&trs*tds&" * from Qianbo_Products where ViewFlag and CommendFlag and Instr(SortPath,'"&SortPath&"')>0 order by id desc"
    'sql="select top "&trs*tds&" * from Qianbo_Products where ViewFlag and NewFlag order by id desc"
    rs.Open sql, conn, 1, 1
    If rs.EOF Then
        response.Write "暂无相关信息"
    Else
        Response.Write "<table cellpadding=""0"" cellspacing=""0"">"&vbCrLf
        For tr = 1 To trs
            Response.Write "  <tr>"&vbCrLf
            For td = 1 To tds
                If StrLen(rs("ProductName"))<= 18 Then
                    ProductName = rs("ProductName")
                Else
                    ProductName = StrLeft(rs("ProductName"), 16)
                End If
                If ISHTML = 1 Then
                    AutoLink = ""&ProName&""&Separated&""&rs("ID")&"."&HTMLName&""
                Else
                    AutoLink = "ProductView.Asp?ID="&rs("id")&""
                End If
                SmallPicPath = HtmlSmallPic(rs("GroupID"), rs("SmallPic"), rs("Exclusive"))
                Response.Write "    <td><table cellpadding=""2"" cellspacing=""0"" >"&vbCrLf
                Response.Write "        <tr>"&vbCrLf
                Response.Write "          <td style=""border: 1px solid #ccc; text-align:center;""><a href="""&SmallPicPath&""" rel=""lightbox"" title="""&rs("ProductName")&"""><img src="""&SmallPicPath&""" alt="""&rs("ProductName")&""" width=""130"" /></a></td>"&vbCrLf
                Response.Write "        </tr>"&vbCrLf
                Response.Write "        <tr>"&vbCrLf
                Response.Write "          <td style=""height:20px; text-align:center;""><a href="""&AutoLink&""" title="""&rs("ProductName")&""">"&ProductName&"</a></td>"&vbCrLf
                Response.Write "        </tr>"&vbCrLf
                Response.Write "        <tr>"&vbCrLf
                Response.Write "          <td style=""height:20px; text-align:center;"">产品型号："&rs("ProductModel")&"</td>"&vbCrLf
                Response.Write "        </tr>"&vbCrLf
                Response.Write "      </table></td>"&vbCrLf
                Response.Write "    <td width=""8""></td>"&vbCrLf
                rs.movenext
                If rs.EOF Then Exit For
            Next
            Response.Write "  </tr>"&vbCrLf
            If rs.EOF Then Exit For
        Next
        Response.Write "</table>"&vbCrLf
    End If
    rs.Close
    Set rs = Nothing
End Function

Function FriendLinks(trs, tds)
    Dim rs, sql, tr, td, ProductName, SmallPicPath
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select top "&trs * tds&" * from Qianbo_FriendLink where ViewFlag order by ID desc"
    rs.Open sql, conn, 1, 1
    If rs.EOF Then
        response.Write "暂无相关信息"
    Else
        
        For tr = 1 To trs
            For td = 1 To tds
                If StrLen(rs("LinkFace"))<= 18 Then
                    LinkFace = rs("LinkFace")
                Else
                    LinkFace = StrLeft(rs("LinkFace"), 16)
                End If
                
                If rs("LinkType") = 0 Then
                    Response.Write "<a href="""&rs("LinkUrl")&""" title="""&rs("LinkName")&""">"&LinkFace&"</a>"
                Else
                    Response.Write "<a href="""&rs("LinkUrl")&""" target=""_blank""><img BORDER=""0"" src="""&rs("LinkFace")&""" alt="""&rs("LinkName")&""" /></a>"
                End If
                
                rs.movenext
                If rs.EOF Then Exit For
            Next
           
            If rs.EOF Then Exit For
        Next
       
    End If
    rs.Close
    Set rs = Nothing
End Function


Function SelPlay(strUrl, strWidth, StrHeight)
    Dim Exts, isExt
    If strUrl <> "" Then
        isExt = LCase(Mid(strUrl, InStrRev(strUrl, ".") + 1))
    Else
        isExt = ""
    End If
    Exts = "mp3,avi,wmv,asf,mov,rm,ra,ram,rmvb,swf"
    If InStr(Exts, isExt) = 0 Then
        Response.Write "<center>视频文件格式错误</center>"
    Else
        Select Case isExt
            Case "mp3", "avi", "wmv", "asf", "mov"
                Response.Write "<embed id=""MediaPlayer"" src="""&strUrl&""" width="""&strWidth&""" height="""&strHeight&""" loop=""true"" autostart=""true"" ShowTracker=""false"" ShowPositionControls=""false""></embed>" 
            Case "mov", "rm", "ra", "ram", "rmvb"
                Response.Write "<object classid=""clsid:cfcdaa03-8be4-11cf-b84b-0020afbbccfa"" id=""video1"" width="""&strWidth&""" height="""&strHeight&""" viewastext>"&vbCrLf
                Response.Write "<param name=""_extentx"" value=""5503"">"&vbCrLf
                Response.Write "<param name=""_extenty"" value=""1588"">"&vbCrLf
                Response.Write "<param name=""autostart"" value=""false"">"&vbCrLf
                Response.Write "<param name=""shuffle"" value=""0"">"&vbCrLf
                Response.Write "<param name=""prefetch"" value=""0"">"&vbCrLf
                Response.Write "<param name=""nolabels"" value=""0"">"&vbCrLf
                Response.Write "<param name=""src"" value="""&strUrl&""">"&vbCrLf
                Response.Write "<param name=""controls"" value=""imagewindow,statusbar,controlpanel"">"&vbCrLf
                Response.Write "<param name=""console"" value=""raplayer"">"&vbCrLf
                Response.Write "<param name=""loop"" value=""0"">"&vbCrLf
                Response.Write "<param name=""numloop"" value=""0"">"&vbCrLf
                Response.Write "<param name=""center"" value=""0"">"&vbCrLf
                Response.Write "<param name=""maintainaspect"" value=""0"">"&vbCrLf
                Response.Write "<param name=""backgroundcolor"" value=""#000000"">"&vbCrLf
                Response.Write "</object>"&vbCrLf
            Case "swf"
                Response.Write "<object codeBase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"" width="""&strWidth&""" height="""&strHeight&""" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"">"&vbCrLf
                Response.Write "<param name=""_cx"" value=""19050"">"&vbCrLf
                Response.Write "<param name=""_cy"" value=""3863"">"&vbCrLf
                Response.Write "<param name=""FlashVars"" value="""">"&vbCrLf
                Response.Write "<param name=""Movie"" value="""&strUrl&""">"&vbCrLf
                Response.Write "<param name=""Src"" value="""&strUrl&""">"&vbCrLf
                Response.Write "<param name=""Play"" value=""-1"">"&vbCrLf
                Response.Write "<param name=""Loop"" value=""-1"">"&vbCrLf
                Response.Write "<param name=""Quality"" value=""High"">"&vbCrLf
                Response.Write "<param name=""SAlign"" value="""">"&vbCrLf
                Response.Write "<param name=""Menu"" value=""0"">"&vbCrLf
                Response.Write "<param name=""Base"" value="""">"&vbCrLf
                Response.Write "<param name=""AllowScriptAccess"" value="""">"&vbCrLf
                Response.Write "<param name=""Scale"" value=""ShowAll"">"&vbCrLf
                Response.Write "<param name=""DeviceFont"" value=""0"">"&vbCrLf
                Response.Write "<param name=""EmbedMovie"" value=""0"">"&vbCrLf
                Response.Write "<param name=""BGColor"" value="""">"&vbCrLf
                Response.Write "<param name=""SWRemote"" value="""">"&vbCrLf
                Response.Write "<param name=""MovieData"" value="""">"&vbCrLf
                Response.Write "<param name=""SeamlessTabbing"" value=""1"">"&vbCrLf
                Response.Write "<param name=""Profile"" value=""0"">"&vbCrLf
                Response.Write "<param name=""ProfileAddress"" value="""">"&vbCrLf
                Response.Write "<param name=""ProfilePort"" value=""0"">"&vbCrLf
                Response.Write "<embed src="""&strUrl&""" width="""&strWidth&""" height="""&strHeight&""" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" menu=""false""></embed>"&vbCrLf
                Response.Write "</object>"&vbCrLf
        End Select
    End If
End Function

Function WebLocation()
    WebLocation = ""
    If request.QueryString("ID") = "" Then
        WebLocation = WebLocation
    ElseIf Not IsNumeric(request.QueryString("ID")) Then
        WebLocation = WebLocation&"读取错误"
    ElseIf conn.Execute("select * from Qianbo_Products Where ViewFlag and  ID="&request.QueryString("ID")).EOF Then
        WebLocation = WebLocation&"读取错误"
    Else
        Dim rs, sql
        Set rs = server.CreateObject("adodb.recordset")
        sql = "select * from Qianbo_Products where ViewFlag and ID="&request.QueryString("ID")
        rs.Open sql, conn, 1, 1
        WebLocation = WebLocation&SortPathTXT("Qianbo_ProductSort", rs("SortID"))
        rs.Close
        Set rs = Nothing
    End If
End Function

Function SortPathTXT(DataFrom, ID)
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "Select * From "&DataFrom&" where ViewFlag and ID="&ID
    rs.Open sql, conn, 1, 1
    If Not rs.EOF Then
        If ISHTML = 1 Then
            AutoLink = ""&ProSortName&""&Separated&""&rs("ID")&""&Separated&"1."&HTMLName&""
        Else
            AutoLink = "ProductList.Asp?SortID="&rs("ID")&""
        End If
        SortPathTXT = SortPathTXT(DataFrom, rs("ParentID"))&" - <a style=""text-decoration:none; color:#151515"" href="""&AutoLink&"""><span>"&rs("SortName")&"</span></a>"
    End If
    rs.Close
    Set rs = Nothing
End Function

Function WebMenu()
    response.Write "<li><a href=""MessageWrite.Asp"">签写留言</a></li>"
    response.Write "<li><a href=""MessageList.Asp"">所有留言</a></li>"
    response.Write "<li><a href=""MemberMessage.Asp"">我的留言</a></li>"
End Function

Function WebLocation()
    WebLocation = "<a href=""MessageList.Asp"">留言反馈</a> - 签写留言"&vbCrLf
End Function

Sub MemInfo()
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from Qianbo_Members where MemName='"&session("MemName")&"'"
    rs.Open sql, conn, 1, 1
    If rs.EOF Then
        response.Write "<center>参数错误</center>"
    Else
        mMemID = rs("ID")
        If Not rs("RealName") = "" Then
            mLinkman = rs("RealName")
        Else
            mLinkman = rs("MemName")
        End If
        mSex = rs("Sex")
        mCompany = rs("Company")
        mAddress = rs("Address")
        mZipCode = rs("ZipCode")
        mTelephone = rs("Telephone")
        mFax = rs("Fax")
        mMobile = rs("Mobile")
        mEmail = rs("Email")
        SecretFlag = "<input type=""checkbox"" name=""SecretFlag"" value=""1"">"
    End If
    rs.Close
    Set rs = Nothing
End Sub

Function WebMenu()
    response.Write "<li><a href=""zp.asp"">招聘信息</a></li>"
    response.Write "<li><a href=""MemberTalent.Asp"">我的应聘</a></li>"
End Function

Function WebLocation()
    WebLocation = "<a href=""zp.asp"">招聘信息</a>"&vbCrLf
End Function

Function WebContent()
    Dim rs, sql
    Dim idCount
    Dim pages
    pages = JobInfo
    Dim pagec
    Dim page
    page = CLng(request("Page"))
    Dim pagenc
    pagenc = 5
    Dim pagenmax
    Dim pagenmin
    Dim pageprevious
    Dim pagenext
    Dim datafrom
    datafrom = "Qianbo_Jobs"
    Dim datawhere
    datawhere = "where ViewFlag  "
    Dim sqlid
    Dim Myself, PATH_INFO, QUERY_STRING
    PATH_INFO = request.servervariables("PATH_INFO")
    QUERY_STRING = request.ServerVariables("QUERY_STRING")'
    If QUERY_STRING = "" Then
        Myself = PATH_INFO & "?"
    ElseIf InStr(PATH_INFO & "?" & QUERY_STRING, "Page=") = 0 Then
        Myself = PATH_INFO & "?" & QUERY_STRING & "&"
    Else
        Myself = Left(PATH_INFO & "?" & QUERY_STRING, InStr(PATH_INFO & "?" & QUERY_STRING, "Page=") -1)
    End If
    Dim taxis
    taxis = "order by id desc "
    Dim i
    sql = "select count(ID) as idCount from ["& datafrom &"]" & datawhere
    Set rs = server.CreateObject("adodb.recordset")
    rs.Open sql, conn, 0, 1
    idCount = rs("idCount")
    If(idcount>0) Then
    If(idcount Mod pages = 0)Then
    pagec = Int(idcount / pages)
Else
    pagec = Int(idcount / pages) + 1
End If
sql = "select id from ["& datafrom &"] " & datawhere & taxis
Set rs = server.CreateObject("adodb.recordset")
rs.Open sql, conn, 1, 1
rs.pagesize = pages
If page < 1 Then page = 1
If page > pagec Then page = pagec
If pagec > 0 Then rs.absolutepage = page
For i = 1 To rs.pagesize
    If rs.EOF Then Exit For
    If(i = 1)Then
    sqlid = rs("id")
Else
    sqlid = sqlid &","&rs("id")
End If
rs.movenext
Next
End If
If(idcount>0 And sqlid<>"") Then
sql = "select * from ["& datafrom &"] where id in("& sqlid &") "&taxis
Set rs = server.CreateObject("adodb.recordset")
rs.Open sql, conn, 0, 1
Response.Write "<ul class=""nl"">"&vbCrLf
While Not rs.EOF
    If ISHTML = 1 Then
        AutoLink = ""&JobNameDiy&""&Separated&""&rs("ID")&"."&HTMLName&""
    Else
        AutoLink = "jobView.asp?ID="&rs("ID")&""
    End If
    Response.Write "<li><span class=""addTime"">"&FormatDate(rs("Addtime"), 13)&"</span>&nbsp;&nbsp;&nbsp;&nbsp;<span class=""className""><a href="""&AutoLink&""">"&rs("JobName")&"</a></span></li>"&vbCrLf
    Response.Write "<li class=""newsLine""></li>"&vbCrLf
    rs.movenext
Wend
Response.Write "</ul>"&vbCrLf
Else
    response.Write "<center>暂无相关信息</center>"
    Exit Function
End If
Response.Write "<table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"&vbCrLf
Response.Write "  <tr height=""35"">"&vbCrLf
Response.Write "    <td align=""center"">"&vbCrLf
Response.Write "共<strong style=""color:red"">"&idcount&"</strong>条记录 页次：<strong style=""color:red"">"&page&"</strong>/"&pagec&" 每页：<strong style=""color:red"">"&pages&"</strong>条记录" & vbCrLf
pagenmin = page - pagenc
pagenmax = page+pagenc
If(pagenmin<1) Then pagenmin = 1
If ISHTML = 1 Then
    If ID = "" Then
        If(page>1) Then response.Write ("<a href="""&JobSortName&""&Separated&"1."&HTMLName&""" title=""回到第一页""><font face=""webdings"" color=""#000000"">9</font></a> ")
    Else
        If(page>1) Then response.Write ("<a href="""&JobSortName&""&Separated&""&ID&""&Separated&"1."&HTMLName&""" title=""回到第一页""><font face=""webdings"" color=""#000000"">9</font></a> ")
    End If
Else
    If(page>1) Then response.Write ("<a href="""& myself &"Page=1"" title=""回到第一页""><font face=""webdings"" color=""#000000"">9</font></a> ")
End If
If page - (pagenc * 2 + 1)<= 0 Then
    pageprevious = 1
Else
    pageprevious = page - (pagenc * 2 + 1)
End If
If ISHTML = 1 Then
    If ID = "" Then
        If(pagenmin>1) Then response.Write ("<a href="""&JobSortName&""&Separated&""&pageprevious&"."&HTMLName&""" title=""第"& pageprevious &"页""><font face=""webdings"" color=""#000000"">3</font></a> ")
    Else
        If(pagenmin>1) Then response.Write ("<a href="""&JobSortName&""&Separated&""&ID&""&Separated&""&pageprevious&"."&HTMLName&""" title=""第"& pageprevious &"页""><font face=""webdings"" color=""#000000"">3</font></a> ")
    End If
Else
    If(pagenmin>1) Then response.Write ("<a href="""& myself &"Page="& pageprevious &""" title=""第"& pageprevious &"页""><font face=""webdings"" color=""#000000"">3</font></a> ")
End If
If(pagenmax>pagec) Then pagenmax = pagec
For i = pagenmin To pagenmax
    If(i = page) Then
    response.Write ("&nbsp;<strong style=""color:red"">"& i &"</strong>&nbsp;")
Else
    If ISHTML = 1 Then
        If ID = "" Then
            response.Write ("[<a href="""&JobSortName&""&Separated&""&i&"."&HTMLName&""">"& i &"</a>]")
        Else
            response.Write ("[<a href="""&JobSortName&""&Separated&""&ID&""&Separated&""&i&"."&HTMLName&""">"& i &"</a>]")
        End If
    Else
        response.Write ("[<a href="""& myself &"Page="&i&""">"& i &"</a>]")
    End If
End If
Next
If page+(pagenc * 2 + 1)>= pagec Then
    pagenext = pagec
Else
    pagenext = page+(pagenc * 2 + 1)
End If
If ISHTML = 1 Then
    If ID = "" Then
        If(pagenmax<pagec) Then response.Write (" <a href="""&JobSortName&""&Separated&""&pagenext&"."&HTMLName&""" title=""跳转到第"&pagenext&"页""><font face=""webdings"" color=""#999999"">:</font></a> ")
        If(page<pagec) Then response.Write (" <a href="""&JobSortName&""&Separated&""&pagec&"."&HTMLName&""" title=""跳转到第"&pagec&"页""><font face=""webdings"" color=""#000000"">:</font></a>")
    Else
        If(pagenmax<pagec) Then response.Write (" <a href="""&JobSortName&""&Separated&""&ID&""&Separated&""&pagenext&"."&HTMLName&""" title=""跳转到第"&pagenext&"页""><font face=""webdings"" color=""#999999"">:</font></a> ")
        If(page<pagec) Then response.Write (" <a href="""&JobSortName&""&Separated&""&ID&""&Separated&""&pagec&"."&HTMLName&""" title=""跳转到第"&pagec&"页""><font face=""webdings"" color=""#000000"">:</font></a>")
    End If
Else
    If(pagenmax<pagec) Then response.Write (" <a href="""& myself &"Page="& pagenext &""" title=""跳转到第"&pagenext&"页""><font face=""webdings"" color=""#999999"">:</font></a> ")
    If(page<pagec) Then response.Write (" <a href="""& myself &"Page="& pagec &""" title=""跳转到第"&pagec&"页""><font face=""webdings"" color=""#000000"">:</font></a>")
End If
Response.Write "    </td>"&vbCrLf
Response.Write "  </tr>"&vbCrLf
Response.Write "</table>"&vbCrLf
rs.Close
Set rs = Nothing
End Function

%>