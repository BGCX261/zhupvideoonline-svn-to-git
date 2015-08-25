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
   <%set rs=server.createobject("adodb.recordset")
   rs.open "select * from baseSet",conn,1,3
   %>

     <div><%=rs("briefInfo")%></div>
	<%
    rs.close
    set rs=nothing
    conn.close
    set conn=nothing
    %>
  </div>
</div>
<!--#include file="foot.asp" -->

