<!--#include file="Inc/syscode.asp" -->
<!--#include file="Inc/TOP.asp"-->
<!--#include file="Inc/eshopcode.asp"-->
<%  
set rsCou = Server.CreateObject("ADODB.Recordset")

if Session("mylogin")<>"1" then
	sqltext="update syswork set querycount=querycount+1 where code='101'"
	conn.execute(sqltext)
	
	rsCou.Open "webcount",conn,3,3
	rsCou.AddNew
	rsCou("ipaddress")=Request.ServerVariables("Remote_HOST") 
	rsCou("logindate")=now()
	rsCou.Update
	rsCou.Close
			
	mylogin=Session("mylogin")
	Session("mylogin")="1"
else
	mylogin=Session("mylogin")	
end if
%>
<%
function cutstr(tempstr,tempwid)
if len(tempstr)>tempwid then
cutstr=left(tempstr,tempwid)&"..."
else
cutstr=tempstr
end if
end function%>
<script src="js/tab.js" type="text/javascript"></script>
<!-- #include file="Inc/Head.asp" -->
<div id="body">
  <table class="bodytable">
    <tr>
      <td class="indexleft">
	  <div class="indexleftbg">
         <!-- #include file="L_login.asp" -->
         <!-- #include file="L_product.asp" -->
		</div>	  </td>
      <td class="indexline"></td>
	  <td class="indexcenter">
	  <div class="indexcenterbg">
	  <table style="width:100%; margin-bottom:12px;">
            <tr>
			  <td class="tabbg">
              <table class="tab_titlebox">
                  <tr>
                    <td>&nbsp;</td>
                    <td class="tab_titleup">新闻公告</td>
                  </tr>
                </table>
              </td>
            </tr>
            <tr>
              <td style="width:350px; vertical-align:top;">
             
			  <div id="tablayera1" style="display: block;" class='indexone'>
			 <% set rsnews=server.createobject("adodb.recordset")
sqltext1="select top 6 *  from news  order by id desc"
rsnews.open sqltext1,conn,1,1
  If rsnews.eof and rsnews.bof then
    response.write "<div>暂无内容</div>"
  else 
  dim q
      q=0
  dim strnewsName
  rsnews.movefirst
  response.write strnewsName & "<ul>"
  do while not rsnews.eof
     response.write strnewsName & "<li><a href='NewsInfo.asp?id=" & rsnews("id") & "'>["& rsnews("time") & "] " & cutstr(rsnews("title"),12) & "</a></li>"
	 q=q+1
	 if q>=6 then exit do
     rsnews.movenext
  loop
  response.write strnewsName & "</ul>"
  rsnews.close
  set rsnews=nothing
end if 
%>
			  </div>
		      </td>
            </tr>
			<tr><td style="text-align:right; background:#F6F6F6;"><a href="News.asp"><img src="img/more.gif" /></a></td>
			</tr>
        </table>
	  <table style="width:100%;">
	            <tr>
			      <td colspan="2" class="tabbg">
			    <table class="tab_titlebox">
                  <tr>
                    <td>&nbsp;</td>
                    <td class="tab_titleup">最新教程</td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td style="width:350px; vertical-align:top;">
				<div class='indexone'>
<% set rsgbook=server.createobject("adodb.recordset")
sqlgbook="select top 6 * from  newbook order by id desc"
rsgbook.open sqlgbook,conn,1,1
  If rsgbook.eof and rsgbook.bof then
    response.write "<div>暂无内容</div>"
  else 
  rsgbook.movefirst
  response.write strnewsName & "<ul>"
  do while not rsgbook.eof
     response.write strnewsName & "<li><a href='willbookInfo.asp?id=" & rsgbook("id") & "title='" &rsgbook("Title")  & "'>" & left(rsgbook("title"),10) & "</a></li>"
     rsgbook.movenext
  loop
  response.write strnewsName & "</ul>"
  rsgbook.close
  set rsgbook=nothing
end if 
%>

</div></td>
              </tr>
              <tr><td style="text-align:right; background:#F6F6F6;"><a href="willbook.asp"><img src="img/more.gif" /></a></td>
			</tr>
          </table>
	  </div>	  </td>
	  <td class="indexline"></td>
      <td class="indexright">
	  <div class="indexrightbg">
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" >
	  <tr>
            <td>
            <div class="indexadvbg">
<%
'打开站点设置连接
sqladv="select * from adv"
Set rs_adv= Server.CreateObject("ADODB.Recordset")
rs_adv.open sqladv,conn,1,1
cebian=ubbcode(rs_adv("cebian")) 
advurl=ubbcode(rs_adv("advurl")) 
advbrief=ubbcode(rs_adv("advbrief")) 
advpic=ubbcode(rs_adv("advpic")) 
if cebian=1 then
%>
<a href="<%=advurl%>" target="_blank" title="<%=advbrief%>"><img src="<%=advpic%>" /></a>
<%
end if
rs_adv.close
%>
</div>
    		</td>
          </tr>
          <!--<tr>
            <td style="background:#ffe25a;"><img src="img/t_zhaoshang.gif" /></td>
          </tr>-->
		  <tr>
            <td class="indextitle">推 荐 视 频 展 示<a href="products.asp"><span style="color:#3aadfd;float:right; cursor:pointer;">更多>>> </span></a> </td>
          </tr>
		  <tr>
		    <td class="indexproduct" style="border-left:#F79700 solid 1px;border-right:#F79700 solid 1px;">
			<%
			Const New_img=8  
set rs_Product=server.createobject("adodb.recordset")
sqltext="select top " & New_img & " * from Product where showMov=False and Passed=True And Elite=True order by UpdateTime desc"
rs_Product.open sqltext,conn,1,1
if not rs_Product.EOF then
%>
<div id='demo' style='overflow:hidden;height:280px;width:100%;'>
  <!--滚动区的高度和宽度-->
    <div id='demo1'>
                <% Do While Not rs_Product.EOF%>
                  <table width="100%" style="border:#CCC solid 1px; ">
                      <tr>
                        <td style="width:120px;"><a href="ArticleShow.asp?ArticleID=<%=rs_Product("articleid")%>"><img src="<%=rs_Product("Uploaddown")%>" width="120" height="90" border="0" /></a></td>
						<td valign="top"><div id="chpname"><a href="ArticleShow.asp?ArticleID=<%=rs_Product("articleid")%>"><%=rs_Product("Title")%></a></div><div class="chpcontent"><%=gotTopic(delHtml(rs_Product("Content")),160)%></div></td>
                      </tr>
                    </table>
<%
rs_Product.MoveNext
Loop
rs_Product.close
%>
    </div>
	<div id=demo2 style="vertical-align:top;"></div>
  </div>
  <%if New_img >3 then%>
<script>
var Picspeed=15
demo2.innerHTML=demo1.innerHTML
function Marquee1(){
if(demo2.offsetHeight-demo.scrollTop<=0)
demo.scrollTop-=demo1.offsetHeight
else{
demo.scrollTop++
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
%>			</td>
		  </tr>
		  <tr><td><table border="0" cellpadding="0" cellspacing="0" width="100%">
            <tr>
              <td class="indexrbleft"></td>
              <td class="indexrbcenter">&nbsp;</td>
              <td class="indexrbright"></td>
            </tr>
          </table></td></tr>
        </table>
	  </div>	  </td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
