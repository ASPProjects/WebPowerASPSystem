   

              <%Folder(1)%>

                
                
                

<%
  Function Folder(id)
    Dim rs, sql, i, ChildCount, FolderType, FolderName, onMouseUp, ListType
    Set rs = server.CreateObject("adodb.recordset")
    sql = "Select * From Qianbo_ProductSort where ParentID="&id&" order by id asc"
    rs.Open sql, conn, 1, 1
    If id = 0 And rs.recordcount = 0 Then
        response.Write ("<center>暂无产品分类</center>")
        Exit Function
    End If
    i = 1
    While Not rs.EOF
        ChildCount = conn.Execute("select count(*) from Qianbo_ProductSort where ParentID="&rs("id"))(0)
        If ChildCount = 0 Then
            If i = rs.recordcount Then
                FolderType = "SortFileEnd"
            Else
                FolderType = "SortFile"
            End If
            FolderName = rs("SortName")
            onMouseUp = ""
        Else
            If i = rs.recordcount Then
                FolderType = "SortEndFolderClose"
                ListType = "SortEndListline"
                onMouseUp = "EndSortChange('a"&rs("id")&"','b"&rs("id")&"');"
            Else
                FolderType = "SortFolderClose"
                ListType = "SortListline"
                onMouseUp = "SortChange('a"&rs("id")&"','b"&rs("id")&"');"
            End If
            FolderName = rs("SortName")
        End If
        datafrom = "Qianbo_ProductSort"
        If ISHTML = 1 Then
            AutoLink = ""&ProSortName&""&Separated&""&rs("ID")&""&Separated&"1."&HTMLName&""
        Else
            AutoLink = "ProductList.Asp?SortID="&rs("ID")&""
        End If 
        response.Write("<li><a href="""&AutoLink&""" title="""&FolderName&""">"&FolderName&"</a></li>")

        If ChildCount>0 Then

End If
rs.movenext
i = i + 1
Wend

rs.Close
Set rs = Nothing
End Function
%>
