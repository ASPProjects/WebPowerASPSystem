<%
Sub HtmlProSort
    Call SiteInfo()
    totalrec = Conn.Execute("Select count(*) from Qianbo_Products Where ViewFlag")(0)
    totalpage = Int(totalrec / ProInfo)
    If (totalpage * ProInfo)<totalrec Then
        totalpage = totalpage+1
    End If
    If totalpage<= 1 Then
        Call htmll("", "", ""&ProSortName&""&Separated&"1."&HTMLName&"", "ProductList.Asp", "Page=", 1, "", "")
    Else
        For i = 1 To totalpage
            Call htmll("", "", ""&ProSortName&""&Separated&""&i&"."&HTMLName&"", "ProductList.Asp", "Page=", i, "", "")
        Next
    End If
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "select * from Qianbo_ProductSort order by ID desc"
    rs.Open sql, conn, 1, 1
    If rs.EOF Then
        Class_Num = 0
    Else
        Class_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            SortPath = conn.Execute("select * from Qianbo_ProductSort Where ViewFlag And ID="&ID)("SortPath")
            totalrec = Conn.Execute("Select count(*) from Qianbo_Products where ViewFlag and SortPath Like '%"&SortPath&"%'")(0)
            totalpage = Int(totalrec / ProInfo)
            If (totalpage * ProInfo)<totalrec Then
                totalpage = totalpage+1
            End If
            If totalpage<= 1 Then
                Call htmll("", "", ""&ProSortName&""&Separated&""&ID&""&Separated&"1."&HTMLName&"", "ProductList.Asp", "SortID=", ID, "Page=", 1)
            Else
                For i = 1 To totalpage
                    Call htmll("", "", ""&ProSortName&""&Separated&""&ID&""&Separated&""&i&"."&HTMLName&"", "ProductList.Asp", "SortID=", ID, "Page=", i)
                Next
            End If
            Response.Write "<script>bar_img.width="&Fix((Class_Num / rs.recordcount) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""成功生成"&Class_Num&"个分类的HTML静态页面。完成比例：" & FormatNumber(Class_Num / rs.recordcount * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Class_Num = Class_Num + 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub HtmlOtherSort
    Call SiteInfo()
    totalrec = Conn.Execute("Select count(*) from Qianbo_Others Where ViewFlag")(0)
    totalpage = Int(totalrec / OtherInfo)
    If (totalpage * OtherInfo)<totalrec Then
        totalpage = totalpage+1
    End If
    If totalpage<= 1 Then
        Call htmll("", "", ""&OtherSortName&""&Separated&"1."&HTMLName&"", "OtherList.Asp", "Page=", 1, "", "")
    Else
        For i = 1 To totalpage
            Call htmll("", "", ""&OtherSortName&""&Separated&""&i&"."&HTMLName&"", "OtherList.Asp", "Page=", i, "", "")
        Next
    End If
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "select * from Qianbo_OthersSort order by ID desc"
    rs.Open sql, conn, 1, 1
    If rs.EOF Then
        Class_Num = 0
    Else
        Class_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            SortPath = conn.Execute("select * from Qianbo_OthersSort Where ViewFlag And ID="&ID)("SortPath")
            totalrec = Conn.Execute("Select count(*) from Qianbo_Others where ViewFlag and SortPath Like '%"&SortPath&"%'")(0)
            totalpage = Int(totalrec / OtherInfo)
            If (totalpage * OtherInfo)<totalrec Then
                totalpage = totalpage+1
            End If
            If totalpage<= 1 Then
                Call htmll("", "", ""&OtherSortName&""&Separated&""&ID&""&Separated&"1."&HTMLName&"", "OtherList.Asp", "SortID=", ID, "Page=", 1)
            Else
                For i = 1 To totalpage
                    Call htmll("", "", ""&OtherSortName&""&Separated&""&ID&""&Separated&""&i&"."&HTMLName&"", "OtherList.Asp", "SortID=", ID, "Page=", i)
                Next
            End If
            Response.Write "<script>bar_img.width="&Fix((Class_Num / rs.recordcount) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""成功生成"&Class_Num&"个分类的HTML静态页面。完成比例：" & FormatNumber(Class_Num / rs.recordcount * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Class_Num = Class_Num + 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub HtmlNewSort
    Call SiteInfo()
    totalrec = Conn.Execute("Select count(*) from Qianbo_News Where ViewFlag")(0)
    totalpage = Int(totalrec / NewInfo)
    If (totalpage * NewInfo)<totalrec Then
        totalpage = totalpage+1
    End If
    If totalpage<= 1 Then
        Call htmll("", "", ""&NewSortName&""&Separated&"1."&HTMLName&"", "NewsList.Asp", "Page=", 1, "", "")
    Else
        For i = 1 To totalpage
            Call htmll("", "", ""&NewSortName&""&Separated&""&i&"."&HTMLName&"", "NewsList.Asp", "Page=", i, "", "")
        Next
    End If
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "select * from Qianbo_NewsSort order by ID desc"
    rs.Open sql, conn, 1, 1
    If rs.EOF Then
        Class_Num = 0
    Else
        Class_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            SortPath = conn.Execute("select * from Qianbo_NewsSort Where ViewFlag And ID="&ID)("SortPath")
            totalrec = Conn.Execute("Select count(*) from Qianbo_News where ViewFlag and SortPath Like '%"&SortPath&"%'")(0)
            totalpage = Int(totalrec / NewInfo)
            If (totalpage * NewInfo)<totalrec Then
                totalpage = totalpage+1
            End If
            If totalpage<= 1 Then
                Call htmll("", "", ""&NewSortName&""&Separated&""&ID&""&Separated&"1."&HTMLName&"", "NewsList.Asp", "SortID=", ID, "Page=", 1)
            Else
                For i = 1 To totalpage
                    Call htmll("", "", ""&NewSortName&""&Separated&""&ID&""&Separated&""&i&"."&HTMLName&"", "NewsList.Asp", "SortID=", ID, "Page=", i)
                Next
            End If
            Response.Write "<script>bar_img.width="&Fix((Class_Num / rs.recordcount) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""成功生成"&Class_Num&"个分类的HTML静态页面。完成比例：" & FormatNumber(Class_Num / rs.recordcount * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Class_Num = Class_Num + 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub HtmlJobSort
    Call SiteInfo()
    totalrec = Conn.Execute("Select count(*) from Qianbo_Jobs Where ViewFlag")(0)
    totalpage = Int(totalrec / JobInfo)
    If (totalpage * JobInfo)<totalrec Then
        totalpage = totalpage+1
    End If
    If totalpage<= 1 Then
        Call htmll("", "", ""&JobSortName&""&Separated&"1."&HTMLName&"", "JobsList.Asp", "Page=", 1, "", "")
    Else
        For i = 1 To totalpage
            Call htmll("", "", ""&JobSortName&""&Separated&""&i&"."&HTMLName&"", "JobsList.Asp", "Page=", i, "", "")
            Response.Write "<script>bar_img.width="&Fix((i / totalpage) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""成功生成"&i&"个分类的HTML静态页面。完成比例：" & FormatNumber(i / totalpage * 100) & """;</script>"
            Response.Flush
        Next
    End If
End Sub

Sub HtmlDownSort
    Call SiteInfo()
    totalrec = Conn.Execute("Select count(*) from Qianbo_Download Where ViewFlag")(0)
    totalpage = Int(totalrec / DownInfo)
    If (totalpage * DownInfo)<totalrec Then
        totalpage = totalpage+1
    End If
    If totalpage<= 1 Then
        Call htmll("", "", ""&DownSortName&""&Separated&"1."&HTMLName&"", "DownList.Asp", "Page=", 1, "", "")
    Else
        For i = 1 To totalpage
            Call htmll("", "", ""&DownSortName&""&Separated&""&i&"."&HTMLName&"", "DownList.Asp", "Page=", i, "", "")
        Next
    End If
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "select * from Qianbo_DownSort order by ID desc"
    rs.Open sql, conn, 1, 1
    If rs.EOF Then
        Class_Num = 0
    Else
        Class_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            SortPath = conn.Execute("select * from Qianbo_DownSort Where ViewFlag And ID="&ID)("SortPath")
            totalrec = Conn.Execute("Select count(*) from Qianbo_Download where ViewFlag and SortPath Like '%"&SortPath&"%'")(0)
            totalpage = Int(totalrec / DownInfo)
            If (totalpage * DownInfo)<totalrec Then
                totalpage = totalpage+1
            End If
            If totalpage<= 1 Then
                Call htmll("", "", ""&DownSortName&""&Separated&""&ID&""&Separated&"1."&HTMLName&"", "DownList.Asp", "SortID=", ID, "Page=", 1)
            Else
                For i = 1 To totalpage
                    Call htmll("", "", ""&DownSortName&""&Separated&""&ID&""&Separated&""&i&"."&HTMLName&"", "DownList.Asp", "SortID=", ID, "Page=", i)
                Next
            End If
            Response.Write "<script>bar_img.width="&Fix((Class_Num / rs.recordcount) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""成功生成"&Class_Num&"个分类的HTML静态页面。完成比例：" & FormatNumber(Class_Num / rs.recordcount * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Class_Num = Class_Num + 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub HtmlPro
    Call SiteInfo()
    totalrec = Conn.Execute("select count(*) from Qianbo_Products where Stock>0")(0)
    sql = "Select * from Qianbo_Products where Stock>0 order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&ProName&""&Separated&""&ID&"."&HTMLName&"", "ProductView.Asp", "ID=", ID, "", "")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub HtmlOther
    Call SiteInfo()
    totalrec = Conn.Execute("select count(*) from Qianbo_Others")(0)
    sql = "Select * from Qianbo_Others order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&OtherName&""&Separated&""&ID&"."&HTMLName&"", "OtherView.Asp", "ID=", ID, "", "")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub HtmlNews
    Call SiteInfo()
    totalrec = Conn.Execute("select count(*) from Qianbo_News")(0)
    sql = "Select * from Qianbo_News order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&NewName&""&Separated&""&ID&"."&HTMLName&"", "NewsView.Asp", "ID=", ID, "", "")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub HtmlJob
    Call SiteInfo()
    totalrec = Conn.Execute("select count(*) from Qianbo_Jobs")(0)
    sql = "Select * from Qianbo_Jobs order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&JobNameDiy&""&Separated&""&ID&"."&HTMLName&"", "JobsView.Asp", "ID=", ID, "", "")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub HtmlInfo
    Call SiteInfo()
    totalrec = Conn.Execute("select count(*) from Qianbo_About")(0)
    sql = "Select * from Qianbo_About order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&AboutNameDiy&""&Separated&""&ID&"."&HTMLName&"", "About.Asp", "ID=", ID, "", "")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub HtmlDown
    Call SiteInfo()
    totalrec = Conn.Execute("select count(*) from Qianbo_Download")(0)
    sql = "Select * from Qianbo_Download order by ID desc"
    Set Rs = Conn.Execute(sql)
    If totalrec = 0 Then
        Detail_Num = 0
    Else
        Detail_Num = 1
        Do While Not rs.EOF
            ID = rs("ID")
            Call htmll("", "", ""&DownNameDiy&""&Separated&""&ID&"."&HTMLName&"", "DownView.Asp", "ID=", ID, "", "")
            Response.Write "<script>bar_img.width="&Fix((Detail_Num / totalrec) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""已成功生成静态页"&Detail_Num&"页，完成比例：" & FormatNumber(Detail_Num / totalrec * 100) & """;</script>"
            Response.Flush
            rs.movenext
            Detail_Num = Detail_Num + 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub HtmlIndex
    Call SiteInfo()
    totalrec = Conn.Execute("Select count(*) from Qianbo_Products Where ViewFlag")(0)
    totalpage = Int(totalrec / ProInfo)
    If (totalpage * ProInfo)<totalrec Then
        totalpage = totalpage+1
    End If
    If totalpage<= 1 Then
        Call htmll("", "", ""&AdvisoryNameDiy&""&Separated&"1."&HTMLName&"", "ProductAdvisory.Asp", "Page=", 1, "", "")
    Else
        For i = 1 To totalpage
            Call htmll("", "", ""&AdvisoryNameDiy&""&Separated&""&i&"."&HTMLName&"", "ProductAdvisory.Asp", "Page=", i, "", "")
            Response.Write "<script>bar_img.width="&Fix((i / totalpage) * 300)&";"
            Response.Write "bar_txt1.innerHTML=""成功生成"&i&"个分类的HTML静态页面。完成比例：" & FormatNumber(i / totalpage * 100) & """;</script>"
            Response.Flush
        Next
    End If
    Call htmll("", "", "Index."&HTMLname&"", "Index.Asp", "", "", "", "")
    Response.Write "<script>bar_img.width="&Fix((1 / 8) * 300)&";bar_txt1.innerHTML=""成功生成首页。完成比例" & FormatNumber(1 / 8 * 100) & """;</script>"
    Response.Flush
    Call htmll("", "", ""&AboutNameDiy&"."&HTMLname&"", "Company.Asp", "", "", "", "")
    Response.Write "<script>bar_img.width="&Fix((2 / 8) * 300)&";bar_txt1.innerHTML=""成功生成“关于我们”静态页面。完成比例：" & FormatNumber(2 / 8 * 100) & """;</script>"
    Response.Flush
    Call htmll("", "", ""&NewSortName&"."&HTMLname&"", "NewsList.Asp", "", "", "", "")
    Response.Write "<script>bar_img.width="&Fix((3 / 8) * 300)&";bar_txt1.innerHTML=""成功生成“新闻列表”静态页面。完成比例：" & FormatNumber(3 / 8 * 100) & """;</script>"
    Response.Flush
    Call htmll("", "", ""&ProSortName&"."&HTMLname&"", "ProductList.Asp", "", "", "", "")
    Response.Write "<script>bar_img.width="&Fix((4 / 8) * 300)&";bar_txt1.innerHTML=""成功生成“产品列表”静态页面。完成比例：" & FormatNumber(4 / 8 * 100) & """;</script>"
    Response.Flush
    Call htmll("", "", ""&JobSortName&"."&HTMLname&"", "JobsList.Asp", "", "", "", "")
    Response.Write "<script>bar_img.width="&Fix((5 / 8) * 300)&";bar_txt1.innerHTML=""成功生成“人才列表”静态页面。完成比例：" & FormatNumber(5 / 8 * 100) & """;</script>"
    Response.Flush
    Call htmll("", "", ""&DownSortName&"."&HTMLname&"", "DownList.Asp", "", "", "", "")
    Response.Write "<script>bar_img.width="&Fix((6 / 8) * 300)&";bar_txt1.innerHTML=""成功生成“下载列表”静态页面。完成比例：" & FormatNumber(6 / 8 * 100) & """;</script>"
    Response.Flush
    Call htmll("", "", ""&OtherSortName&"."&HTMLname&"", "OtherList.Asp", "", "", "", "")
    Response.Write "<script>bar_img.width="&Fix((7 / 8) * 300)&";bar_txt1.innerHTML=""成功生成“其他信息列表”静态页面。完成比例：" & FormatNumber(7 / 8 * 100) & """;</script>"
    Response.Flush
    conn.Close
    Set conn = Nothing
End Sub
%>