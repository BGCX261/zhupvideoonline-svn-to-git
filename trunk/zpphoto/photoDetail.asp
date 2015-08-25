<!--#include file="dbConn.asp"-->
<!--#include file="head.asp"-->
<%
   dim sql
   dim rs
   dim listname
   set rs=server.createobject("adodb.recordset")
   sql="update albums set hits=hits+1 where articleID="&request("id")
   rs.open sql,conn,1,3
   sql="select * from albums where articleid="&request("id")
   rs.open sql,conn,1,1
%>
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
<div class="pCont">
<!--#include file="top.asp"-->
<div class="main">
      <div id="plane1" style="background:#6699cc;filter:alpha(opacity=100)">图片载入中,请稍候...</div>
      <div id="plane0" style="display:none;" class="pPhoto">
        <div class="pTitle"><%=rs("title")%> </div>
              <div> 作者：<%=rs("author")%> &nbsp; 来源：<%=rs("source")%> &nbsp; 上传/更新：<%=rs("dateandtime")%> &nbsp; 拍摄时期：<%=rs("orgTime")%> &nbsp; 观看次数：<%=rs("hits")%> </div>
              <div class="pContent"><%=rs("content")%></div>
			  <%
  if rs("lx")=0 then
  %>
              <%if rs("images1")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=1" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images2")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=2" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images3")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=3" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images4")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=4" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images5")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=5" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images6")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=6" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images7")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=7" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images8")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=8" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images9")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=9" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images10")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=10" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images11")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=11" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images12")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=12" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images13")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=13" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images14")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=14" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images15")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=15" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images16")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=16" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images17")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=17" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images18")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=18" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images19")<>"" then%>
              <div class="smlpicborder"><img src="pics.asp?id=<%=request("id")%>&pics=19" onerror="this.src='IMAGES/nopic.gif'" /></div>
              <%end if %>
              <%if rs("images20")<>"" then%>
              <div class="smlpicborder"><a href="javascript:winopen('photopic.asp?ID=<%=rs("articleid")%>&pic=20')"> <img src="pics.asp?id=<%=request("id")%>&pics=20" onerror="this.src='IMAGES/nopic.gif'" /></a></div>
              <%end if %>
              <%
  elseif rs("lx")=1 then
   %>
              <div class="smlpicborder"><object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0' width='640' height='480' border='0'>
                <param name=movie value='<%= rs("image1") %>'>
                <param name=quality value=high>
                <embed src='<%= rs("image1") %>' quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='640' height='480'></embed>
              </object></div>
              <%
  end if
%>
</div>
<div class="pinfo">			<% rs.close                 
set rs=server.createobject("adodb.recordset")                 
sql="select * from albums where articleid="&request("id")+1                 
rs.open sql,conn,1,1                 
  if rs.eof and rs.bof then                 
                  %>
              |&lt;上一篇:已经没有了
              <%else%>
              &lt;&lt;上一篇:<a href="photoDetail.asp?id=<%=rs("articleid")%>"><%=rs("title")%> </a>
              <%end if%>
			  
			<%rs.close 
set rs=server.createobject("adodb.recordset") 
sql="select * from albums where articleid="&request("id")-1 
rs.open sql,conn,1,1  
  if rs.eof and rs.bof then 
                  %>
              &gt;|下一篇:已经没有了
              <%else%>
              &gt;&gt;下一篇:<a href="photoDetail.asp?id=<%=rs("articleid")%>"><%=rs("title")%> </a>
              <%end if%>           </div>
</div>
</div>
<%
rs.close
set rs=nothing
%>
<!--#include file="foot.asp" -->

<script>
setTimeout("plane1.style.display='none';")
setTimeout("plane0.style.display='block';")//显示页面内容
</script>