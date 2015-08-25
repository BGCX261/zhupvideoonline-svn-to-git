<%
Const New_img=8  
set rs_Product=server.createobject("adodb.recordset")
sqltext="select top " & New_img & " *  from Product where Passed=True And Elite=True order by UpdateTime desc"
rs_Product.open sqltext,conn,1,1
if not rs_Product.EOF then%>
<div align='center' id='demo' style='overflow:hidden;height:125px;width:500px;'>
  <!--滚动区的高度和宽度-->
  <table align='center' cellpadding='0' cellspace='0' border='0'>
    <tr>
      <td id='demo1' valign='top'>
                <table>
                  <tr align="center">
                    <% Do While Not rs_Product.EOF%>
                    <td><table border="1" bordercolor="#eeeeee" style="border-collapse:collapse;">
                        <tr>
                          <td align="center"><a href="ArticleShow.asp?ArticleID=<%=rs_Product("articleid")%>"><img src="<%=rs_Product("DefaultPicUrl")%>" height="80" /></a></td>
                        </tr>
                        <tr>
                          <td align="center" bgcolor="#eeeeee"><a href="ArticleShow.asp?ArticleID=<%=rs_Product("articleid")%>"><%=rs_Product("Title")%></a></td>
                        </tr>
                      </table></td>
                    <%
rs_Product.MoveNext
Loop
rs_Product.close
%>
                  </tr>
                </table>
                </td>
      <td id=demo2 valign=top></td>
    </tr>
  </table>
</div>
<%if New_img >4 then%>
<script>
var Picspeed=15
demo2.innerHTML=demo1.innerHTML
function Marquee1(){
if(demo2.offsetWidth-demo.scrollLeft<=0)
demo.scrollLeft-=demo1.offsetWidth
else{
demo.scrollLeft++
}
}
var MyMar1=setInterval(Marquee1,Picspeed)
demo.onmouseover=function() {clearInterval(MyMar1)}
demo.onmouseout=function() {MyMar1=setInterval(Marquee1,Picspeed)}
</script>
	<%end if
else
	Response.Write "暂无数据"
end if
rs_Product.close
set rs_Product=nothing
%>