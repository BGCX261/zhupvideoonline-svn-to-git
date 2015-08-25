
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<meta name="keywords" content="<%=keywords%>" />
<meta name="description" content="<%=describe%>" />
<title><%if page_title<>"" then%><%=page_title%> - <%end if%><%if channel_title<>"" then%><%=channel_title%> - <%end if%><%=site_title%></title>
<link href="<%=S_TPL_PATH%>css.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="<%=S_TPL_PATH%>js/script.js"></script>
<!--[if IE 6]>
<script type="text/javascript" src="<%=S_TPL_PATH%>js/DD_belatedPNG_0.0.8a.js"></script>
<script type="text/javascript">
DD_belatedPNG.fix("*");
</script>
<![endif]-->
</head>
<body>
  
<div id="header">
  <div class="logo"><a href="<%=S_ROOT%>"><img src="<%=S_ROOT%>images/logo.jpg" /></a></div>
  <div class="search">
    <form method="post" action="<%=S_ROOT%>index.asp?search.<%=suffix%>">
      <input class="text" name="search" type="text" maxlength="30" />
      <%if S_LANG<>"en-us" then%>
      <input class="button" name="submit" type="submit" value="<%=L_SEARCH%>"/>
      <%else%>
      <input class="button" name="submit" type="submit" style="width:35px;font-size:10px;font-weight:normal;" value="<%=L_SEARCH%>"/>
      <%end if%>
    </form>
  </div>
  <div id="nav">
    <ul>
      <%for i=0 to nav_u%>
      <li <%if i=0 then%>class="first"<%end if%>>
	  	<a href="<%=nav(i,1)%>" <%if nav(i,2)=1 then%>target="_blank"<%end if%>><%=nav(i,0)%></a>
		
		<%word=nav(i,0)%>
		<%if word="关于我们" then%>
			
		<%elseif word="产品展示" then%>
			
		<%elseif word="资讯中心" then%>
			
		<%end if%>
		
		
	  	<%if i=1 then%>
			<ul>
				<%for j=0 to nav_about_u%>
				<li><a href="<%=nav_about(j,1)%>"><%=nav_about(j,0)%></a></li>
				<%next%>
			</ul>
		<%elseif i=2 then%>
			<ul>
			  <%for j=0 to nav_pro_u%>
				<%if nav_pro(j,3)=1 then%>
				<li><a href="<%=S_ROOT%>?product_<%=nav_pro(j,0)%>_0_0.<%=suffix%>"><%=nav_pro(j,2)%></a></li>
				<%end if%>
			  <%next%>
			</ul>		
		<%elseif i=3 then%>
			<ul>
			  <%for j=0 to nav_art_u%>
				<%if nav_art(j,3)=1 then%>
				<li><a href="<%=S_ROOT%>?article_<%=nav_art(j,0)%>_0_0.<%=suffix%>"><%=nav_art(j,2)%></a></li>
				<%end if%>
			  <%next%>
			</ul>		
		<%end if%>
	  </li>
	  <%next%>
      <div class="clear"></div>
    </ul>
  </div>
</div>
<!-- 新秀 -->

  
<div id="focus">
  <%if i>=0 then%>
  <div id="focus_bg"></div>
  <div id="focus_show"></div>
  <div id="focus_img">
  <%for i=0 to focus_u%>
    <div name="focus_img" id="focus_<%=i+1%>"><%=focus(i,0)%>|<%=focus(i,1)%>|<%=focus(i,2)%></div>
  <%next%>
  </div>
  <script type="text/javascript" src="<%=S_TPL_PATH%>js/focus.js"></script>
  <%end if%>
</div>
<!-- 新秀 -->

  <div id="main">
    <div id="left">
      
<div class="box tree">
  <div class="head"><span><%=L_ABOUT_MENU%></span></div>
  <div class="main">
  <%for i=0 to menu_u%>
    <div class="cat1"><a href="<%=menu(i,1)%>"><%=menu(i,0)%></a></div>
  <%next%>
  </div>
</div>
<!-- 新秀 -->

    </div>
    <div id="right">
      
<div class="here">
  <span><%if page_title<>"" then%><%=cut_str(page_title,10)%><%else%><%=channel_title%><%end if%></span>
  <div class="link">
  <a href="<%=S_ROOT%>"><%=L_HOME%></a>
  <%if channel_title<>"" then%><img src="<%=S_TPL_PATH%>images/here_img.jpg" /><a href="<%=S_ROOT%>?<%=url_page%>.<%=suffix%>"><%=channel_title%></a><%end if%>
  <%if page_title<>"" then%><img src="<%=S_TPL_PATH%>images/here_img.jpg" /><a href="<%=S_ROOT%>?<%=url%>"><%=cut_str(page_title,20)%></a><%end if%>
  </div>
</div>
<div id="about_main">
  <%if about_u>=0 then%><%=about(0,1)%><%end if%>
</div>
<!-- 新秀 -->

    </div>
    <div class="clear"></div>
  </div>
  
<div id="footer">
  <div class="nav">
  <%for i=0 to footer_nav_u%>
    <a href="<%=footer_nav(i,1)%>"><%=footer_nav(i,0)%></a>
    <%if i<>footer_nav_u then%>|<%end if%>
  <%next%>
  </div>
  <div class="info">
  Powered by <a href="http://www.sinsiu.com/" target="_blank">sinsiu</a>
  <a href="<%=sit_code_url%>" target="_blank"><%=sit_code%></a>
  <a href="<%=sit_tec_url%>" target="_blank"><%=sit_tec%></a>
  <%=sit_count_code%>
  </div>
</div>
<!-- 新秀 -->

  
<div id="service">
  <div id="service_img" onmousemove="show_service()"></div>
  <div class="main">
    <div class="in">
      <%=service%>
    </div>
  </div>
</div>
<script type="text/javascript" src="<%=S_TPL_PATH%>js/service.js"></script>
<!-- 新秀 -->
</body>
</html>