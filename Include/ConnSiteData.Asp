<%
Dim Conn, ConnStr, DataBase
Set Conn = Server.CreateObject("Adodb.Connection")
Set FSO = Server.CreateObject("Scripting.FileSystemObject")
If FSO.FileExists(Server.MapPath("./Database/"&SiteDataPath&"")) Then
    DataBase = Server.MapPath("./Database/"&SiteDataPath&"")
ElseIf FSO.FileExists(Server.MapPath("../Database/"&SiteDataPath&"")) Then
    DataBase = Server.MapPath("../Database/"&SiteDataPath&"")
End If
ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & DataBase
Conn.Open ConnStr
If Err Then
    Err.Clear
    Set Conn = Nothing
    Response.Write "数据库连接错误，请检查连接参数！"
    Response.End
End If
%>
<!--#include file="Function.Asp" -->