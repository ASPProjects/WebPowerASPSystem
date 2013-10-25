<!--#include file="../part/header.asp"-->
<img src="template/SweeLife/image/banner_sweelife.jpg" />
<script type="text/javascript">
var imgUrl1="template/gufeng/image/banner_1.jpg";
var imgtext1="广东佛山市景秀装饰材料有限公司";
var imgLink1=escape("#");
var imgUrl2="template/gufeng/image/banner_2.jpg";
var imgtext2="广东佛山市景秀装饰材料有限公司";
var imgLink2=escape("#");
var imgUrl3="template/gufeng/image/banner_3.jpg";
var imgtext3="广东佛山市景秀装饰材料有限公司";
var imgLink3=escape("#");

var focus_width=1024;
var focus_height=494;
var text_height=0;
var swf_height = focus_height+text_height;

var pics=imgUrl1+"|"+imgUrl2+"|"+imgUrl3;
var links=imgLink1+"|"+imgLink2+"|"+imgLink3;
var texts=imgtext1+"|"+imgtext2+"|"+imgtext3;
var movie = 'static/image/focus2.swf';
var objText =''+
'<object '+
'width="'+ focus_width +'" '+
'height="'+ swf_height +'" '+
'classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" '+
'codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" ' +
'style="vertical-align:middle;">'+
'<param name="movie" value="'+movie+'?pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'" />'+
'<param name="quality" value="high" />'+
'<param name="wmode" value="opaque" />'+
'<embed '+
'src="'+movie+'?pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'" '+
'width="'+ focus_width +'" '+
'height="'+ swf_height +'" '+
'quality="high" '+
'wmode="opaque" '+
'pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" '+
'type="application/x-shockwave-flash" style="vertical-align:middle;"></embed>'+
'</object>';
document.write(objText);
//document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ focus_width +'" height="'+ swf_height +'">');
//document.write('<param name="allowScriptAccess" value="sameDomain"><param name="movie" value="images/focus2.swf"><param name="quality" value="high"><param name="bgcolor" value="#F0F0F0">');
//document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
//document.write('<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'">');
//document.write('</object>');
</script>
<div id="products">

<div id="SW">
<a href="#"><div><span></span></div>SWN多彩世界系列</a>
<a href="#"><div><span></span></div>SWW可户外玉白系列</a>
</div>

<script src="Scripts/ScrollPic.js" type="text/javascript"></script>
<div id="ProductsShow">
<ul id="ISL_Cont_2" style="margin: 3px 60px 0;" class="cl">
<li><a href="#"><img src="static/image/SWN001.jpg" /></a></li>
<li><a href="#"><img src="static/image/SWN002.jpg" /></a></li>
<li><a href="#"><img src="static/image/SWN003.jpg" /></a></li>
<li><a href="#"><img src="static/image/SWN004.jpg" /></a></li>
<li><a href="#"><img src="static/image/SWN005.jpg" /></a></li>
<li><a href="#"><img src="static/image/SWN006.jpg" /></a></li>
</ul>
<div id="LeftArr" style="
cursor: pointer;
width: 60px;
height: 100%;
position: absolute;
top: 0;"></div>
<div id="RightArr" style="
cursor: pointer;
width: 60px;
height: 100%;
position: absolute;
top: 0;
right:0;"></div>
</div>
<script type="text/javascript" src="Scripts/magnifier.js"></script>
<script type="text/javascript">
<!--//--><![CDATA[//><!--
var scrollPic_01 = new ScrollPic();
scrollPic_01.scrollContId = "ISL_Cont_2"; //内容容器ID
scrollPic_01.arrLeftId = "LeftArr";//左箭头ID
scrollPic_01.arrRightId = "RightArr"; //右箭头ID

scrollPic_01.frameWidth = jQuery("#ISL_Cont_2").css('width');//显示框宽度
scrollPic_01.pageWidth = 99; //翻页宽度

scrollPic_01.speed = 10; //移动速度(单位毫秒，越小越快)
scrollPic_01.space = 10; //每次移动像素(单位px，越大越快)
scrollPic_01.autoPlay = true; //自动播放
scrollPic_01.autoPlayTime = 3; //自动播放间隔时间(秒)

scrollPic_01.initialize(); //初始化

//--><!]]>
</script>
<script type="text/javascript">
magnifier("ProductsShow");
</script>
</div>
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