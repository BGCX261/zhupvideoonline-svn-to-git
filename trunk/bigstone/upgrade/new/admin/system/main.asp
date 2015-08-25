<!-- #include file="common.asp" -->
<%
dim dir:dir = gets("dir","")
dim fil:fil = gets("fil","")
if dir = "" then dir = "basic"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<title><%=L_PAGE_TITLE%></title>
<link href="../system/css.css" rel="stylesheet" type="text/css">
<style type="text/css">html,body{overflow-x:hidden;overflow-y:hidden;}</style>
<script type="text/javascript" src="../system/script.min.js"></script>
</head>
<body>
<div id="header">
  <ul id="nav">
	<%
	dim nav_admin:nav_admin = get_nav("admin",0)
	for i = 0 to ubound(nav_admin)
	%>
    <li><a href="<%=nav_admin(i,2)%>"><%=nav_admin(i,1)%></a></li>
    <%next%>
  </ul>
  <div class="link">
    <a href="<%=S_ROOT%>" target="_blank">网站首页</a>
    <a href="?out=1">退出系统</a>
  </div>
  <%call interface("admin_header","")%>
</div>
<div id="left">
  <ul>
	<%
    dim nav_type:nav_type = "admin_" & dir
    if dir = "about" or dir = "recruit" or dir = "download" then nav_type = "admin_article"
    dim nav_arr:nav_arr = get_nav(nav_type,0)
    for i = 0 to ubound(nav_arr)
    %>
    <li><a href="<%=nav_arr(i,2)%>"><%=nav_arr(i,1)%></a></li>
    <%next%>
  </ul>
  <div id="version">
    Powered by <a href="http://www.sinsiu.com/" target="_blank">sinsiu</a><br>
    version:<%=get_config("version")%>
  </div>
</div>
<div id="right">
  <%if fil <> "" then fil = fil & ".asp" else fil = "basic_info.asp"%>
  <iframe frameborder="no" src="<%="../"&dir&"/"&fil%>?module=<%=dir%>"></iframe>
</div>
<div style="clear:both"></div>
<div id="footer"></div>
<script>var ie6 = false;</script>
<!--[if IE 6]><script>ie6 = true;</script><![endif]-->
<script language="javascript">
  width = document.getElementById("footer").offsetLeft;
  height = document.getElementById("footer").offsetTop;
  document.getElementById("left").style.height = height + "px";
  document.getElementById("right").style.width = width - 183 + "px";
  if(ie6)
  {
	  document.getElementById("right").style.height = height - 38 + "px";
  }else{
	  document.getElementById("right").style.height = height - 53 + "px";
  }
</script>
</body>
</html>
<%
call db.close_db()
set db = nothing
'新秀
%>