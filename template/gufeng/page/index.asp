<!--#include file="../part/header.asp"-->
<!--#include file="../part/productsclass.asp"-->
<script type="text/javascript" src="static/js/imgcrop.js"></script>
<div id="ProductsShow">
<span class="title">产品展示</span>
<div class="splitline"></div>
<a class="ol more" href="product.asp">more</a>
<!--<ul class="cl" style="display: none;">
<li><a href="#"><img src="template/gufeng/image/product1.jpg"/><span>产品展示</span></a> </li>
<li><a href="#"><img src="template/gufeng/image/product1.jpg"/><span>产品展示</span></a> </li>
<li><a href="#"><img src="template/gufeng/image/product1.jpg"/><span>产品展示</span></a> </li>
<li><a href="#"><img src="template/gufeng/image/product1.jpg"/><span>产品展示</span></a> </li>
</ul>-->
<!--#include file="../part/cpgundongcaifeng.asp"-->
</div>
<script type="text/javascript">ImgCrops('#ProductsShow')</script>
<div id="AboutUs">
<span class="title">关于我们</span>
<div class="splitline"></div>
<a class="ol more" href="aboutus.asp">more</a>
<div id="youku">
<img src="template/gufeng/image/tmp_youku.jpg"/>
</div>
<div class="intro">
<p>佛山市固峰自动化机械设备厂，是一家专业从事自动化机械研发、设计、生产、销售和服务为一体的企业，为了响应市场的需求，我们研发出了各领域的自动化机械设备，主要包括：全自动攻牙、钻孔、切管、倒角机械，广泛应用于五金加工跟管材加工领域，企业经过长期的努力，已建立了完善的管理体系，拥有一支高素质的研发团队和技术人才，我们立志为广大客户提供先进、实用、稳定的自动化机械设备，目前销量覆盖面广，畅销全国各地，部分已出口至东南亚、非洲及等国家和地区。我们本着&ldquo;以质量求生存，以服务求发展&rdquo;的企业宗旨与广大新老客户紧密合作，携手共进，共创未来。</p>

</div>
</div>
<div id="ContactUs">
<ul class="Title">
<li>联系人：</li>
<li>手机：</li>
<li>电话：</li>
<li>传真：</li>
<li>邮箱：</li>
<li>网址：</li>
<li>地址：</li>
</ul>

<ul class="Info">
<li>颜生</li>
<li>15986139858 18676528240</li>
<li>0757-23330849</li>
<li>0757-23330849</li>
<li>64540657@qq.com</li>
<li>www.fsgfjx.com</li>
<li>广东省佛山市顺德区陈村镇石洲岗北工业区一号</li>
</ul>
</div>
<div class="banner"></div>
<div class="wp_spc"><div id="QRCode"></div></div>
<div class="splitline" style="clear: left; margin:0 5px;"></div>
<div id="Strength">
<span class="title">实力展示</span>
<div class="splitline"></div>
<a class="ol more" href="#">more</a>
<ul class="cl"><%=Products("1,",1,6)%>
<!--<li><a href="#"><img src="template/gufeng/image/product_small1.jpg"/><span>钻孔机</span></a> </li>
<li><a href="#"><img src="template/gufeng/image/product_small1.jpg"/><span>钻孔机</span></a> </li>
<li><a href="#"><img src="template/gufeng/image/product_small1.jpg"/><span>钻孔机</span></a> </li>
<li><a href="#"><img src="template/gufeng/image/product_small1.jpg"/><span>钻孔机</span></a> </li>
<li><a href="#"><img src="template/gufeng/image/product_small1.jpg"/><span>钻孔机</span></a> </li>
<li><a href="#"><img src="template/gufeng/image/product_small1.jpg"/><span>钻孔机</span></a> </li>-->
</ul>
</div>
<div id="Success">
<span class="title">成功案例</span>
<div class="splitline"></div>
<a class="ol more">more</a>
<div id="Cases">
<a class="ol"></a>
<a class="ol"></a>
<a class="ol"></a>
<a class="ol"></a>
<a class="ol"></a>
<a class="ol"></a>
<a class="ol"></a>
<a class="ol"></a>
<a class="ol"></a>
<a class="ol"></a>
<a class="ol"></a>
<a class="ol"></a>
</div>
</div>
<div id="News">
<span class="title">新闻中心</span>
<div class="splitline"></div>
<div class="switch">
<span class="act">行业新闻</span><span>企业新闻</span>
</div>
<a class="ol more" href="news.asp">more</a>
<ul id="news_company" class="cl">
<%=News("5,")%>
</ul>
<ul id="news_Industry" class="cl">
<%=News("1,")%>
</ul>
</div>
<script type="text/javascript" src="static/js/switch.js"></script>
<script type="text/javascript">
bindSwitch('News',{'企业新闻':'news_company','行业新闻':'news_Industry'})
</script>
<!--#include file="../part/links.asp"-->
<!--#include file="../part/footer.asp"-->
<%
Function WebMenu(ParentID, i, level)
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from Qianbo_NewsSort where ViewFlag and ParentID="&ParentID&" order by ID asc"
    rs.Open sql, conn, 1, 1
    If conn.Execute("select ID from Qianbo_NewsSort Where ViewFlag and ParentID=0").EOF Then
        response.Write "<center>暂无相关信息</center>"
    End If
    Do While Not rs.EOF
        If ISHTML = 1 Then
            AutoLink = ""&NewSortName&""&Separated&""&rs("ID")&""&Separated&"1."&HTMLName&""
        Else
            AutoLink = "NewsList.Asp?SortID="&rs("ID")&""
        End If
        'response.Write "<li><a href="""&AutoLink&""">"&rs("SortName")&"</a></li>"
        If rs("SortName") = "行业新闻" Then
            cla = " class=""act"""
        Else
            cla = ""
        End If
        Response.Write("<span"&cla&">"&rs("SortName")&"</span>")
        i = i + 1
        If i<level Then Call WebMenu(rs("ID"), i, level)
        i = i -1
        rs.movenext
    Loop
    rs.Close
    Set rs = Nothing
End Function

Function WebLocation()
    WebLocation = "<a href=""NewsList.Asp"">公司动态</a>"&vbCrLf
    If request.QueryString("SortID") = "" Then
        WebLocation = WebLocation
    ElseIf Not IsNumeric(request.QueryString("SortID")) Then
        WebLocation = WebLocation&"参数错误"
    ElseIf conn.Execute("select * from Qianbo_NewsSort Where ViewFlag and  ID="&request.QueryString("SortID")).EOF Then
        WebLocation = WebLocation&"参数错误"
    Else
        Dim rs, sql
        Set rs = server.CreateObject("adodb.recordset")
        sql = "select * from Qianbo_NewsSort where ViewFlag and ID="&request.QueryString("SortID")
        rs.Open sql, conn, 1, 1
        WebLocation = WebLocation&SortPathTXT("Qianbo_NewsSort", rs("ID"))
        rs.Close
        Set rs = Nothing
    End If
End Function

Function ChildSort()
    Dim ParentID
    ParentID = request.QueryString("SortID")
    If ParentID = "" Or (Not IsNumeric(ParentID)) Then Exit Function
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from Qianbo_NewsSort where ViewFlag and ParentID="&ParentID&" order by ID desc"
    rs.Open sql, conn, 1, 1
    If rs.bof And rs.EOF Then
        Exit Function
    Else
        While Not rs.EOF
            If ISHTML = 1 Then
                AutoLink = ""&NewSortName&""&Separated&""&rs("ID")&""&Separated&"1."&HTMLName&""
            Else
                AutoLink = "NewsList.Asp?SortID="&rs("ID")&""
            End If
            response.Write "<a href="""&AutoLink&""">"&rs("SortName")&"</a> | "
            rs.movenext
        Wend
    End If
    rs.Close
    Set rs = Nothing
End Function

Function SortPathTXT(DataFrom, ID)
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "Select * From "&DataFrom&" where ViewFlag and ID="&ID
    rs.Open sql, conn, 1, 1
    If Not rs.EOF Then
        If ISHTML = 1 Then
            AutoLink = ""&NewSortName&""&Separated&""&rs("ID")&""&Separated&"1."&HTMLName&""
        Else
            AutoLink = "NewsList.Asp?SortID="&rs("ID")&""
        End If
        SortPathTXT = SortPathTXT(DataFrom, rs("ParentID"))&" - <a href="""&AutoLink&""">"&rs("SortName")&"</a>"
    End If
    rs.Close
    Set rs = Nothing
End Function

Function WebContent(DataFrom, ID, SortPath)
    Dim rs, sql
    Dim HideSort
    Set rs = server.CreateObject("adodb.recordset")
    If ID = "" Then
        SortPath = "0,"
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
    pages = NewInfo
    Dim pagec
    Dim page
    page = CLng(request("Page"))
    Dim pagenc
    pagenc = 5
    Dim pagenmax
    Dim pagenmin
    Dim pageprevious
    Dim pagenext
    datafrom = "Qianbo_News"
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
        AutoLink = ""&NewName&""&Separated&""&rs("ID")&"."&HTMLName&""
    Else
        AutoLink = "NewsView.Asp?ID="&rs("ID")&""
    End If
    Response.Write "<li><span class=""addTime"">"&FormatDate(rs("Addtime"), 13)&"</span><span class=""className""><a href="""&AutoLink&""">"&rs("NewsName")&"</a></span></li>"&vbCrLf
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
        If(page>1) Then response.Write ("<a href="""&NewSortName&""&Separated&"1."&HTMLName&""" title=""回到第一页""><font face=""webdings"" color=""#000000"">9</font></a> ")
    Else
        If(page>1) Then response.Write ("<a href="""&NewSortName&""&Separated&""&ID&""&Separated&"1."&HTMLName&""" title=""回到第一页""><font face=""webdings"" color=""#000000"">9</font></a> ")
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
        If(pagenmin>1) Then response.Write ("<a href="""&NewSortName&""&Separated&""&pageprevious&"."&HTMLName&""" title=""第"& pageprevious &"页""><font face=""webdings"" color=""#000000"">3</font></a> ")
    Else
        If(pagenmin>1) Then response.Write ("<a href="""&NewSortName&""&Separated&""&ID&""&Separated&""&pageprevious&"."&HTMLName&""" title=""第"& pageprevious &"页""><font face=""webdings"" color=""#000000"">3</font></a> ")
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
            response.Write ("[<a href="""&NewSortName&""&Separated&""&i&"."&HTMLName&""">"& i &"</a>]")
        Else
            response.Write ("[<a href="""&NewSortName&""&Separated&""&ID&""&Separated&""&i&"."&HTMLName&""">"& i &"</a>]")
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
        If(pagenmax<pagec) Then response.Write (" <a href="""&NewSortName&""&Separated&""&pagenext&"."&HTMLName&""" title=""跳转到第"&pagenext&"页""><font face=""webdings"" color=""#999999"">:</font></a> ")
        If(page<pagec) Then response.Write (" <a href="""&NewSortName&""&Separated&""&pagec&"."&HTMLName&""" title=""跳转到第"&pagec&"页""><font face=""webdings"" color=""#000000"">:</font></a>")
    Else
        If(pagenmax<pagec) Then response.Write (" <a href="""&NewSortName&""&Separated&""&ID&""&Separated&""&pagenext&"."&HTMLName&""" title=""跳转到第"&pagenext&"页""><font face=""webdings"" color=""#999999"">:</font></a> ")
        If(page<pagec) Then response.Write (" <a href="""&NewSortName&""&Separated&""&ID&""&Separated&""&pagec&"."&HTMLName&""" title=""跳转到第"&pagec&"页""><font face=""webdings"" color=""#000000"">:</font></a>")
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