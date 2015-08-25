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
<html>
<head>
<title>welcome!</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta http-equiv="refresh" content="15;URL=cn_index.asp">
<style>
body {FONT-SIZE: 12px; BACKGROUND:#c0c0c0 url(img/bg.gif);Margin:0;}
td {FONT-SIZE: 12px;}
img{ border:0;}
A              {COLOR:#00f;text-decoration:none;}
A:link,visited {COLOR:#00f;text-decoration:none;}  
A:hover        {COLOR:#f00;FONT-SIZE:12px;text-decoration: underline} 
.fontdiv       {font-size:12px; text-align:center; color:#ccc;}
</style>
</head>
<body>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center">
      <div style=" margin:2px;">
        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="400" height="325" title="学习宝贝">
          <param name="movie" value="img/index.swf">
          <param name="quality" value="high">
          <embed src="img/index.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="400" height="325"></embed>
        </object>
      </div> 
	  <div class="fontdiv">如不能正常播放动画,请下载安装相关浏览器flash播放插件.<a href="Cn_index.asp">如需跳过请点击此处>>></a></div>
	       </td>
  </tr>
</table>
</body>
</html>