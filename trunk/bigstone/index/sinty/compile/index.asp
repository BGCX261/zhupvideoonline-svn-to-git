
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<meta name="keywords" content="<%=keywords%>" />
<meta name="description" content="<%=describe%>" />
<title><%=site_title%></title>
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
      
<div class="box" id="notice">
  <div class="head"><span><%=L_NOTICE%></span></div>
  <div class="main">
  <%=notice%>
  </div>
</div>
<!-- 新秀 -->

      
<div class="box" id="contact">
  <div class="head"><span><%=L_CONTACT%></span></div>
  <div class="main">
    <%for i=0 to contact_u%>
    <span><%=contact(i,0)%>：</span><%=contact(i,1)%><br />
    <%next%>
  </div>
</div>
<!-- 新秀 -->

      
<div class="box" id="link">
  <div class="head"><span><%=L_LINK%></span></div>
  <div class="main">
    <%if img_link_u>=0 then%>
    <div class="img">
      <%for i=0 to img_link_u%>
      <a href="<%=img_link(i,0)%>" title="<%=img_link(i,2)%>" target="_blank"><img src="<%=img_link(i,1)%>"/></a>
      <%next%>
      <div class="clear"></div>
    </div>
    <%end if%>
    <%if word_link_u>=0 then%>
    <div class="word">
      <%for i=0 to word_link_u%>
      <a href="<%=word_link(i,0)%>" title="<%=word_link(i,2)%>" target="_blank"><%=word_link(i,1)%></a>
      <%next%>
    </div>
    <%end if%>
  </div>
</div>
<!-- 新秀 -->

    </div>
    <div id="right">
      
<div class="box2" id="about">
  <div class="head"><span><%if about_title>="" then%><%=about_title%><%else%><%=L_ABOUT_US%><%end if%></span><a class="more" href="<%=S_ROOT%>?about.<%=suffix%>"><%=L_MORE%></a></div>
  <div class="main">
    <div class="img"><img src="<%=S_TPL_PATH%>images/about.jpg" /></div>
    <%=about_text%>
    <a class="more" href="<%=S_ROOT%>?about.<%=suffix%>">【<%=L_DETAILED%>】</a>
  </div>
</div>
<!-- 新秀 -->

      
<div class="box2" id="recommend">
  <div class="head"><span><%=L_RECOMMEND%></span><a class="more" href="<%=S_ROOT%>?product.<%=suffix%>"><%=L_MORE%></a></div>
  <div class="main">
  <!-------------------------->
  <div id="roll">
    <div id="roll_sheet" onmouseover="over_roll()" onmouseout="out_roll()">
      <%for i=0 to recommend_u%>
      <div name="roll_unit">
        <table cellspacing="0" cellpadding="0">
          <tr><td class="img">
            <a href="<%=S_ROOT%>?product_<%=recommend(i,0)%>.<%=suffix%>" target="_blank"><img src="<%=S_ROOT%><%=recommend(i,2)%>" onload="picresize(this,150,150)"/></a>
          </td></tr>
          <tr><td class="title">
            <a href="<%=S_ROOT%>?product_<%=recommend(i,0)%>.<%=suffix%>" target="_blank"><%=cut_str(recommend(i,1),11)%></a>
          </td></tr>
        </table>
      </div>
      <%next%>
    </div>
  </div>
  <!-------------------------->
  </div>
</div>
<script type="text/javascript" src="<%=S_TPL_PATH%>js/roll.js"></script>
<!-- 新秀 -->

      
<div class="box2 list" id="news">
  <div class="head"><span><%=L_NEWS%></span><a class="more" href="<%=S_ROOT%>?article.<%=suffix%>"><%=L_MORE%></a></div>
  <ul class="main">
    <%for i=0 to news_u%>
    <li><a href="<%=S_ROOT%>?article_<%=news(i,0)%>.<%=suffix%>" target="_blank"><%=cut_str(news(i,1),20)%></a><span><%=datevalue(news(i,2))%></span><div class="clear"></div></li>
    <%next%>
  </ul>
</div>
<!-- 新秀 -->

      
<div class="box2 list" id="product">
  <div class="head"><span><%=L_NEW_PRODUCT%></span><a class="more" href="<%=S_ROOT%>?product.<%=suffix%>"><%=L_MORE%></a></div>
  <ul class="main">
    <%for i=0 to product_u%>
	<li><a href="<%=S_ROOT%>?product_<%=product(i,0)%>.<%=suffix%>" target="_blank"><%=cut_str(product(i,1),20)%></a><span><%=datevalue(product(i,2))%></span><div class="clear"></div></li>
    <%next%>
  </ul>
</div>
<!-- 新秀 -->

      <div class="clear"></div>
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