<div id="demo" style="width:445px; overflow:hidden; margin:0px auto">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td id="demo1"><table border="0" cellspacing="0" cellpadding="0">
          <tr>
           
                
                
                
                  <%
	   
    Set rs = server.CreateObject("adodb.recordset")
        sql = "Select  top 8 * From Qianbo_Products where ViewFlag  and ID in(101,112,114,111,107,108)"
    rs.Open sql, conn, 1, 1
	 IF Not rs.eof Then  
					while Not rs.Eof 
	%>
    
          
 <td><table width="120" border="0" cellspacing="0" cellpadding="0">
                <tr>
                
                
                  <td valign="top" background="images/indpicbg.jpg"><table width="120" border="0" cellspacing="0" cellpadding="0" style="margin:5px">
                      <tr>
                        <td height="5"></td>
                      </tr>
                      <tr>
                        <td align="center" style="border:1px #CCCCCC solid;" ><a href="ProductView.Asp?ID=<% =rs("ID")%>" target=_blank><img style=" padding:2px" src="<% =rs("SmallPic")%>" height="1230px" width="180px"  border="0" title="<% =rs("ProductName")%>"></a></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td height="20" align="center"><a href="ProductView.Asp?ID=<% =rs("ID")%>" style="font-size:12PX; color:#797940; text-decoration:none" target=_blank><% =left(rs("ProductName"),8)%></a></td>
                </tr>
              </table></td>
 
         <%
		
					   	 rs.MoveNext
      Wend
	End IF 
		
		%>  
 
 
          </tr>
          
          <tr>
           
                
                
                
                  <%
	   
    Set rs = server.CreateObject("adodb.recordset")
        sql = "Select  top 8 * From Qianbo_Products where ViewFlag  and ID in(98,96,104,105,109,110)"
    rs.Open sql, conn, 1, 1
	 IF Not rs.eof Then  
					while Not rs.Eof 
	%>
    
          
 <td><table width="120" border="0" cellspacing="0" cellpadding="0">
                <tr>
                
                
                  <td valign="top" background="images/indpicbg.jpg"><table width="120" border="0" cellspacing="0" cellpadding="0" style="margin:5px">
                      <tr>
                        <td height="5"></td>
                      </tr>
                      <tr>
                        <td align="center" style="border:1px #CCCCCC solid;" ><a href="ProductView.Asp?ID=<% =rs("ID")%>" target=_blank><img style=" padding:2px" src="<% =rs("SmallPic")%>" height="130px" width="180px"  border="0" title="<% =rs("ProductName")%>"></a></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td height="20" align="center"><a href="ProductView.Asp?ID=<% =rs("ID")%>" style="font-size:12PX; color:#797940; text-decoration:none" target=_blank><% =left(rs("ProductName"),8)%></a></td>
                </tr>
              </table></td>
 
         <%
		
					   	 rs.MoveNext
      Wend
	End IF 
		
		%>  
 
 
          </tr>
             
          
          
          
          <tr> </tr>
        </table></td>
      <td id="demo2"></td>
    </tr>
  </table>
</div>
<script>
var speed1=20
demo2.innerHTML=demo1.innerHTML
function Marquee1(){
if(demo2.offsetWidth-demo.scrollLeft<=0)
demo.scrollLeft=0
else{
demo.scrollLeft++
}
}
var MyMar1=setInterval(Marquee1,speed1)
demo.onmouseover=function() {clearInterval(MyMar1)}
demo.onmouseout=function() {MyMar1=setInterval(Marquee1,speed1)}
</script>
