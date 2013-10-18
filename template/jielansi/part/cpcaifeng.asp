
<%=WebContent("Qianbo_ProductSort",request.QueryString("SortID") ,"")%>
<%
   
   Function WebContent(DataFrom, ID, SortPath)
    Dim rs, sql
    Dim HideSort
    Set rs = server.CreateObject("adodb.recordset")
    If ID = "" Then 
        SortPath = "0,1,2,"
    ElseIf Not IsNumeric(ID) Then
        response.Write "<center>暂无相关信息</center>"
        Exit Function
    ElseIf conn.Execute("select * from "&DataFrom&" Where ViewFlag and  ID="&ID).EOF Then
        response.Write "<center>暂无相关信息</center>"
        Exit Function
    Else
        SortPath = conn.Execute("select * from "&DataFrom&" Where ViewFlag and  ID="&ID)("SortPath")
        conn.Execute("update "&DataFrom&" set ClickNumber=ClickNumber+1 Where ID="&ID)
    End If
    sql = "select * from "&DataFrom&" Where not(ViewFlag) and Instr(SortPath,'"&SortPath&"')>0"
    rs.Open sql, conn, 1, 1
    While Not rs.EOF
        HideSort = "and not(Instr(SortPath,'"&rs("SortPath")&"')>0) "&HideSort
        rs.movenext
    Wend
    rs.Close
    Dim idCount
    Dim pages
    pages = ProInfo
    Dim pagec
    Dim page
    page = CLng(request("Page"))
    Dim pagenc
    pagenc = 4
    Dim pagenmax
    Dim pagenmin
    Dim pageprevious
    Dim pagenext
    datafrom = "Qianbo_Products"
    Dim datawhere
    datawhere = "where ViewFlag and Instr(SortPath,'"&SortPath&"')>0 "&HideSort& " "
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
    taxis = "order by id asc "
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
rs.Open sql, conn, 1, 1






    OtherPic = rs("OtherPic")
    If Not(IsNull(OtherPic)) Then
        OtherPic = Split(OtherPic, "*")
    End If





Dim tr, td
Dim ProductName, SmallPicPath, Content
Response.Write "<ul class=""cpzs"">"&vbCrLf
For tr = 1 To 6

    For td = 1 To 6
        If StrLen(rs("ProductName"))<= 12 Then
            ProductName = rs("ProductName")
        Else
            ProductName = StrLeft(rs("ProductName"), 30)
        End If
        If ISHTML = 1 Then
            AutoLink = ""&ProName&""&Separated&""&rs("ID")&"."&HTMLName&""
        Else
            AutoLink = "ProductView.Asp?ID="&rs("ID")&"&SortID="&rs("sortID")&"&SortPath="&rs("SortPath")&""
        End If
        SmallPicPath = HtmlSmallPic(rs("GroupID"), rs("SmallPic"), rs("Exclusive"))
		
		Response.Write "<li><a href="&AutoLink&" title="""&rs("ProductName")&""" ><img onload=""tupianjiaz(this)"" src="""&SmallPicPath&""" alt="&rs("ProductName")&" width=""188px"" height=""152""></a><a class=""name"" href="&AutoLink&""&request.QueryString("SortID")&">"&ProductName&"</a></li>"
		
       
        rs.movenext
		
        If rs.EOF Then Exit For
		
    Next

    If rs.EOF Then Exit For
Next


Response.Write "</ul>"&vbCrLf


Else
    response.Write "<center>暂无相关信息</center>"
    Exit Function
End If






Response.Write "<div class=""yema"">"&vbCrLf
Response.Write "共"&idcount&"条记录 页次："&page&"</strong>/"&pagec&" 每页："&pages&"条记录" & vbCrLf
pagenmin = page - pagenc
pagenmax = page+pagenc
If(pagenmin<1) Then pagenmin = 1
If ISHTML = 1 Then
    If ID = "" Then
        If(page>1) Then response.Write ("<a href="""&ProSortName&""&Separated&"1."&HTMLName&""" title=""回到第一页""><font face=""webdings"" color=""#000000"">9</font></a> ")
    Else
        If(page>1) Then response.Write ("<a href="""&ProSortName&""&Separated&""&ID&""&Separated&"1."&HTMLName&""" title=""回到第一页""><font face=""webdings"" color=""#000000"">9</font></a> ")
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
        If(pagenmin>1) Then response.Write ("<a href="""&ProSortName&""&Separated&""&pageprevious&"."&HTMLName&""" title=""第"& pageprevious &"页""><font face=""webdings"" color=""#000000"">3</font></a> ")
    Else
        If(pagenmin>1) Then response.Write ("<a href="""&ProSortName&""&Separated&""&ID&""&Separated&""&pageprevious&"."&HTMLName&""" title=""第"& pageprevious &"页""><font face=""webdings"" color=""#000000"">3</font></a> ")
    End If
Else
    If(pagenmin>1) Then response.Write ("<a href="""& myself &"Page="& pageprevious &""" title=""第"& pageprevious &"页""><font face=""webdings"" color=""#000000"">3</font></a> ")
End If
dim xianshi

If(pagenmax>pagec) Then pagenmax = pagec


For i = pagenmin To pagenmax
    If(i = page) Then
    'response.Write ("&nbsp;<strong style=""color:red"">"& i &"</strong>&nbsp;")
	xianshi=i
Else
'    If ISHTML = 1 Then
'        If ID = "" Then
'            response.Write ("[<a href="""&ProSortName&""&Separated&""&i&"."&HTMLName&""">"& i &"</a>]")
'        Else
'            response.Write ("[<a href="""&ProSortName&""&Separated&""&ID&""&Separated&""&i&"."&HTMLName&""">"& i &"</a>]")
'        End If
'    Else
'        response.Write ("[<a href="""& myself &"Page="&i&""">"& i &"</a>]")
'    End If
End If
Next

if xianshi>1 then

response.Write ("[<a href="""& myself &"Page="&(xianshi-1)&""">上一页</a>]")
else 
end if

if xianshi<pagec then

response.Write("[<a href="""& myself &"Page="&(xianshi+1)&""">下一页</a>]")

end if

If page+(pagenc * 2 + 1)>= pagec Then
    pagenext = pagec
Else
    pagenext = page+(pagenc * 2 + 1)
End If
If ISHTML = 1 Then
    If ID = "" Then
        If(pagenmax<pagec) Then response.Write (" <a href="""&ProSortName&""&Separated&""&pagenext&"."&HTMLName&""" title=""跳转到第"&pagenext&"页"">跳转到第"&pagenext&"页</a> ")
        If(page<pagec) Then response.Write (" <a href="""&ProSortName&""&Separated&""&pagec&"."&HTMLName&""" title=""跳转到第"&pagec&"页""><font face=""webdings"" color=""#000000"">:</font></a>")
    Else
        If(pagenmax<pagec) Then response.Write (" <a href="""&ProSortName&""&Separated&""&ID&""&Separated&""&pagenext&"."&HTMLName&""" title=""跳转到第"&pagenext&"页""><font face=""webdings"" color=""#999999"">:</font></a> ")
        If(page<pagec) Then response.Write (" <a href="""&ProSortName&""&Separated&""&ID&""&Separated&""&pagec&"."&HTMLName&""" title=""跳转到第"&pagec&"页""><font face=""webdings"" color=""#000000"">:</font></a>")
    End If
Else
   ' If(pagenmax<pagec) Then response.Write (" <a href="""& myself &"Page="& pagenext &""" title=""跳转到第"&pagenext&"页"">跳转到第"&i&"页</a> ")
    If(page<pagec) Then response.Write (" <a href="""& myself &"Page="& pagec &""" title=""跳转到第"&pagec&"页""><font face=""webdings"" color=""#000000"">:</font></a>")
End If
Response.Write "</div>"&vbCrLf
rs.Close
Set rs = Nothing
End Function
   %>
