<!--#include file="Inc/syscode.asp" -->
<!--#include file="Inc/TOP.asp"-->
<!--#include file="Inc/eshopcode.asp"-->
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
		<div align="left">
        <img src="img/2.gif" width="178" />
        <%set rspro=server.createobject("adodb.recordset")
sqlpro="select top 8 *  from ProSet  order by mdy_time desc"
rspro.open sqlpro,conn,1,1
  If rspro.eof and rspro.bof then
    response.write "暂无内容"
  else 
  dim strnewsName
  rspro.movefirst
  response.write strnewsName & "<ul>"
  do while not rspro.eof
response.write strnewsName & "<li id='one'><a href='"& rspro("siteUrl") & "'  target='_blank'>" & cutstr(rspro("title"),14) & "</a></li>"
     rspro.movenext
  loop
  response.write strnewsName & "</ul>"
  rspro.close
  set rspro=nothing
end if
%>
        <img src="img/1.gif" width="178" />
        <% 
		sqltype="select * from RecruitType Order by id asc"
		Set rstype= Server.CreateObject("ADODB.Recordset")
		rstype.open sqltype,conn,1,1
	if rstype.bof and rstype.eof then 
		response.Write "没有数据"
	else
		rstype.movefirst
		response.write "<ul>"
		do while not rstype.eof
			response.write "<li id='one'><a href='recruit.asp?typed="&rstype("typed")&"'>"&rstype("typed")&"</a>"
			response.write "</li>"
			rstype.movenext
		loop
		response.write "</ul>"
		rstype.close
		set rstype=nothing
	end if
 %>
        <img src="img/3.gif" width="178" />
	<%  set rslk=server.createobject("adodb.recordset")
	sqllk = "select top 10 * from links order by id desc"
	rslk.open sqllk,conn,1,1%>
	  <select onChange="var optlk=this.options[selectedIndex].value; if(optlk!='0') window.open(optlk)" style="width:100%;">
	<option selected value="">友情链接</option>
<%
if not (rslk.bof and rslk.eof) then
	rslk.movefirst
	do while not rslk.eof
        response.Write "<option value='" & trim(rslk("link")) & "'>" & trim(rslk("name")) & "</option>"
	   	rslk.movenext
	loop
end if
%>
</select>          </div></div>	  </td>
      <td class="indexline"></td>
	  <td class="indexcenter">
			  <div class="tab_titledown"><div class="indexFloatLeft"> 通知公告</div><div class="indexFloatRight"><a href="News.asp">更多>>></a></div>
	      <!--#include file="HotNews.asp" --></div>
	      <div class="tab_titledown"><div class="indexFloatLeft"> 学院新闻</div><div class="indexFloatRight"><a href="News.asp">更多>>></a></div>
          <ul> 
				<%
t=0
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="SELECT * from news where typed='学院新闻' order by id desc,mdy_time desc" 
rs.Open sql,conn,1,1
if not Rs.eof then
do while not rs.eof
t=t+1
%>
                    
                      <li  <%if rs("toTop")=true then response.write "class='toTop'" else response.write "class=''" end if%>><a href="NewsInfo.asp?id=<%=rs("id")%>" title="<%=rs("title")%>" target="_blank"><%=cutstr(rs("title"),24)%></a><span class="noticeTime"><%=DateValue(rs("mdy_time"))%></span></li>
                    
					<% 
if t>=4 then exit do 
rs.movenext 
loop 
else 
response.write "<div align=center>暂无新闻!</div>" 
end if 
rs.close 
%>
  </ul></div>
	  </td>
	  <td class="indexline"></td>
      <td class="indexright">
<div class="indexadvbg">
<%
'打开站点设置连接
sqladv="select * from adv"
Set rs_adv= Server.CreateObject("ADODB.Recordset")
rs_adv.open sqladv,conn,1,1
cebian=rs_adv("cebian") 
advurl=rs_adv("advurl") 
advbrief=rs_adv("advbrief") 
advpic=rs_adv("advpic")
advfloat=rs_adv("advfloat")
imgFloat=rs_adv("imgFloat")
lnkFloat=rs_adv("lnkFloat") 
titleFloat=rs_adv("titleFloat")
if cebian=1 then
%>
<a href="<%=advurl%>" target="_blank" title="<%=advbrief%>"><img src="<%=advpic%>" border=0 /></a>
<%
end if
rs_adv.close
%>
</div>
       <!-- #include file="L_stuZone.asp" -->
       <div class="tab_titledown"><div class="indexFloatLeft">网络课程推荐</div><div class="indexFloatRight"><a href="article.asp">更多>>></a></div>
              <div class="indexart">
              <%
			Const New_img=3  
set rs_Product=server.createobject("adodb.recordset")
sqltext="select top " & New_img & " * from Product where Passed=True And Elite=True order by UpdateTime desc"
rs_Product.open sqltext,conn,1,1
if not rs_Product.EOF then
%>
                <% Do While Not rs_Product.EOF%>
                  <div id="indexartname"><a href="ArticleShow.asp?ArticleID=<%=rs_Product("articleid")%>"><img src="img/arrow3.gif">  <%=rs_Product("Title")%></a></div><div class="indexartcontent"><%=gotTopic(delHtml(rs_Product("Content")),90)%></div>
<%
rs_Product.MoveNext
Loop
rs_Product.close
set rs_Product=nothing
%>


<%end if


%>
              </div></div>
      </td>
    </tr>
    <tr>
      <td colspan="5">
        <div class="movlayer">
<%
Const stu_img=8 
set rs_stu=server.createobject("adodb.recordset")
sqlstu="select top " & stu_img & " * from StuZone where Elite=True  And  Passed=True order by UpdateTime desc"
rs_stu.open sqlstu,conn,1,1
if not rs_stu.EOF then
%>
<div align='center' id='demo' style='overflow:hidden;height:150px;width:880px;'>
              <!--滚动区的高度和宽度-->
              <table align='center' cellpadding='0' cellspace='0' border='0'>
                <tr>
                  <td id='demo1' valign='top'>
              <table width="100%">
                <tr align="center">
                  <% Do While Not rs_stu.EOF%>
                  <td><table width="100%" title="<%=rs_stu("Title")%>">
                      <tr>
                        <td width="35%" align="center" style="border:#CCC solid 1px;"><a href="stuZoneInfo.asp?id=<%=rs_stu("id")%>"><img src="<%=rs_stu("DefaultPicUrl")%>" height="120" border="0" /></a>
						<div id="chpname"><a href="stuZoneInfo.asp?id=<%=rs_stu("id")%>"><%=cutstr(rs_stu("Title"),6)%></a></div></td>
                      </tr>
                    </table></td>
                  <%
rs_stu.MoveNext
Loop
rs_stu.close
%>
                </tr>
              </table>			  </td>
                  <td id=demo2 valign=top></td>
                </tr>
              </table>
			  </div>
<%if stu_img >4 then%>
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
	Response.Write "无数据"
end if
rs_stu.close
set rs_stu=nothing
%>	
</div>
      </td>
    </tr>
  </table>
</div>
 <script language="javascript" type="text/javascript">
            // JavaScript Document
            //设置广告块 
            var divLeft = 0;
            //设置广告块初始左边位置 
            var divTop = 0;
            //设置广告块初始顶部位置 
            var divWidth = 80;
            //设置广告块宽度 
            var divHeight = 80;
            //设置广告块高度 
            var divImg = "<%=imgFloat%>";
            //设置广告图片的URL地址 
            var divUrl = "<%=lnkFloat%>";
            //设置广告链接 
            var divTitle = "<%=titleFloat%>";
            //设置div标题
            document.write("<div id=\"adDiv\" style=\"position:absolute; left:"+divLeft);
            document.write("px; top:"+divTop+"px; width:"+divWidth+"px; height:"+divHeight);
            document.write("px; z-index:1;\" onMouseOver=\"javascript:window.clearInterval(varId)\"");
            document.write(" onMouseOut=\"javascript:beginMoveFAd();\"><a href=\""+divUrl+"\" target=\"_blank\">");
            document.write("<img src=\""+divImg+"\" border=\"0\" title=\""+divTitle+"\"></a><br>");
            document.write("<a href=\"javascript:\" onclick=\"closeFAd();\">");
            document.write("<font color=\"#000000\">关闭广告</font></a>");
            document.write("</div>");
            //飘浮广告 
            var _stepx=2;_stepy=2;
            //初始化每次偏离的象素 
            var moveSpeed=40;
            //速度 
            var varId;
            //获取setInterval的ID 
            function moveFAd()
            {
             //飘浮广告主函数
             var adLeft=parseInt(adDiv.style.left);
             var adTop=parseInt(adDiv.style.top);
             var adWidth=parseInt(adDiv.style.width);
             var adHeight=parseInt(adDiv.style.height);
             var _bodyLeft=document.body.scrollLeft;
             var _bodyTop=document.body.scrollTop;
             var _bodyHeight=document.body.clientHeight+_bodyTop;
             var _bodyWidth=document.body.clientWidth+_bodyLeft;
             if(adLeft<=_bodyLeft)
             {
              _stepx=2;
             }
             if(adTop<=_bodyTop)
             {
              _stepy=2;
             } 
             if((adLeft+adWidth)>=_bodyWidth)
             {
              _stepx=-2;
             }
             if((adTop+adHeight)>=_bodyHeight)
             {
              _stepy=-2;
             }
             adDiv.style.left=adLeft+_stepx;
             adDiv.style.top=adTop+_stepy;
            }
            function beginMoveFAd()            {
             //启动飘浮
             varId = window.setInterval("moveFAd()",moveSpeed);
            }
			function closeFAd(){
			 document.getElementById("adDiv").style.display="none";
			}
             
			var floatAd = "<%=advfloat%>";
			if(floatAd == 1){
				//给窗体的load事件附加飘浮函数
                //window.onload=beginMoveFAd;
				
				function addLoadEvent(func){		
					var oldonload=window.onload;	
					if(typeof window.onload!='function'){	
						window.onload=func;	
					}else{
						window.onload=function(){
							oldonload();
                            func();	
						}
					}	
				}
				addLoadEvent(beginMoveFAd);
				/**/
				
			}else{
				closeFAd();
			}
    </script>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
