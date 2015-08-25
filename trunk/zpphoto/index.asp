<!--#include file="dbConn.asp"-->
<!--#include file="head.asp"-->
<body>
<script type="text/javascript">
 var Sys = {};
 var ua = navigator.userAgent.toLowerCase();
 var s;
 (s = ua.match(/msie ([\d.]+)/)) ? Sys.ie = s[1] :
 (s = ua.match(/firefox\/([\d.]+)/)) ? Sys.firefox = s[1] :
 (s = ua.match(/chrome\/([\d.]+)/)) ? Sys.chrome = s[1] :
 (s = ua.match(/opera.([\d.]+)/)) ? Sys.opera = s[1] :
 (s = ua.match(/version\/([\d.]+).*safari/)) ? Sys.safari = s[1] : 0;

//以下进行测试
 if (Sys.ie) document.write('<div style="top:2px;left:0;position:absolute; width:100%; height:100%;">');
 if (Sys.firefox) document.write('<div style="top:2px;position:absolute; width:100%; height:100%;z-index:-1;">');
 if (Sys.chrome) document.write('<div style="top:2px;position:absolute; width:100%; height:100%;z-index:-1;">');
 if (Sys.opera) document.write('<div style="top:2px;position:absolute; width:100%; height:100%;z-index:-1;">');
 if (Sys.safari) document.write('<div style="top:2px;position:absolute; width:100%; height:100%;z-index:-1;">');
</script>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="100%" height="100%">
              <param name="movie" value="images/spider.swf" />
              <param name="quality" value="low" />
              <param name="wmode" value="transparent" />
              <param name="SCALE" value="exactfit">
              <embed src="images/spider.swf" width="100%" type="application/x-shockwave-flash" height="100%" quality="low" pluginspage="http://www.macromedia.com/go/getflashplayer" wmode="transparent" scale="exactfit"></embed>
     
</object>
</div>
<div class="cont">
<!--#include file="top.asp"-->
  <div class="main">
    <div class="left">
        <ul class="paihang">
          <li class="paihangtitle">:: 浏览排行榜 </li>
          <%
Set rs1= Server.CreateObject("ADODB.Recordset")
sql ="SELECT dateandtime,hits,articleid,title FROM albums WHERE now()-dateandtime <=180 ORDER BY hits DESC"
rs1.open sql,conn,1,1
                dim w
                w=0
                do while not rs1.eof%>
          <li><a title="加入时间：<%=rs1("dateandtime")%>" href="photoDetail.asp?id=<%=rs1("articleid")%>"><%=rs1("title")%></a></li>
          <%w=w+1
	 if w>=6 then exit do
	    rs1.movenext 
	   loop 
	   rs1.close
	   set rs1=nothing
	   		        %>
        </ul>
      </div>
    <div class="right">
        <div class="alm">注意:网站可能访问较慢请耐心等待。图片有时可能会出现问题，如遇图片不能自适应缩小等情况请多刷新几遍。</div>
		<div class="rtitle">最新图集</div>
          <ul class="txtlist">
            
            <%
sql="SELECT type.typeid,type.type,albums.typeid,albums.articleid,albums.title,albums.dateandtime FROM albums,type where type.typeid=albums.typeid ORDER BY albums.dateandtime DESC"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
dim i
i=0
do while not rs.eof 
                        %>
            <li> [<a href="photoType.asp?typeid=<%=rs("typeid")%>" target="_top" class="blue"><%=rs("type")%></a>]<a title="<%=rs("title")%>" href="photoDetail.asp?id=<%=rs("articleid")%>" target="_blank">
              <%note=rs("title")
            if len(note)>12 then note=left(note,12)&"..."
            response.write note
                    %>
              </a><%=(Year(rs("dateandtime")) & "/" & String(2-Len(month(rs("dateandtime"))),"0") & month(rs("dateandtime")) & "/" & String(2-Len(day(rs("dateandtime"))),"0") & day(rs("dateandtime")))%> </li>
            <%i=i+1
	if i>9 then exit do
	rs.movenext
	loop
	rs.close
	set rs=nothing
                        %>
          </ul>
        
        <div class="rtitle">推荐图集</div>
        <%set rst=server.createobject("adodb.recordset")
                  dim y
                  y=1
    sql="select articleid,bestpic,title,lx,images1 from albums where best=true order by articleid desc"             
rst.open sql,conn,1,1
do while not rs.eof
y=y+1
%>
        <dl class="smlpicbox">
          <dt class="smlpic">
            <%
  if rst("lx")=0 then
  %>
            <a href='photoDetail.asp?id=<%= rst("articleid") %>' target='_blank'><img src='bestpics.asp?id=<%= rst("articleid") %>' onerror=""this.src='IMAGES/nopic.gif'"" / ></a>
            <%
  elseif rst("lx")=1 then
   %>
            <a href='photoDetail.asp?id=<%= rst("articleid") %>' target='_blank'> <img src='bestpics.asp?id=<%= rst("articleid") %>' onerror=""this.src='IMAGES/nopic.gif'"" / > </a>
            <%
  end if
%>
          </dt>
          <dd class="pictitle"><%=rst("title")%></dd>
        </dl>
        <%
                    if y>5 then exit do
                    rst.movenext
                    loop                                                      
 rst.close
 set rst=nothing
                  %>
           </div>
  </div>
</div>
<!--#include file="foot.asp" -->