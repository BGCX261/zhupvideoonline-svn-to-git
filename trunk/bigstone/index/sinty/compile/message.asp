
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<meta name ="keywords" content="<%=keywords%>" />
<meta name ="description" content="<%=describe%>" />
<title><%=channel_title%> - <%=site_title%></title>
<link href="<%=S_TPL_PATH%>css.css" rel="stylesheet" type="text/css">
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
  <div class="main">
    <div class="logo"><a href="<%=S_ROOT%>"><img src="<%=S_ROOT%>images/logo.png" /></a></div>
    <div class="search">
      <form name="form_search" method="post" action="<%=S_ROOT%>index.asp?search.<%=suffix%>">
        <input class="text" name="search" type="text" maxlength="30" />
        <%if S_LANG<>"en-us" then%>
        <a class="button" onclick="document.form_search.submit();"><%=L_SEARCH%></a>
        <%else%>
        <a class="button" onclick="document.form_search.submit();" style="font-size:10px;font-weight:normal;"><%=L_SEARCH%></a>
        <%end if%>
      </form>
    </div>
  </div>
  <div id="nav">
    <ul>
      <%for i=0 to nav_u%>
      <li <%if i=0 then%>class="first"<%end if%>><a href="<%=nav(i,1)%>" <%if nav(i,2)=1 then%>target="_blank"<%end if%>><%=nav(i,0)%></a></li>
      <%next%>
      <div class="clear"></div>
    </ul>
  </div>
</div>
<!-- 新秀 -->

  <div id="main">
    
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

    <div id="left">
      
<div class="box" id="research">
  <div class="head"><span><%=L_RESEARCH%></span></div>
  <div class="main">
  <%if research_u>=0 then%>
    <form action="<%=S_ROOT%>index/submit.asp" method="post">
    <input name="cmd" type="hidden" value="research"/>
    <%sign=chr(123)&"v"&chr(125)%>
    <%k=0%>
    <%for i=0 to research_u%>
      <%arr=split(research(i,1),sign)%>
      <div class="question"><%=arr(0)%></div>
      <div>
      <%for j=2 to ubound(arr)%>
        <%select case arr(1)%>
          <%case "radio":%><span><input name="res_<%=k%>" type="radio" value="<%=arr(j)%>" /><%=arr(j)%></span>
          <%case "checkbox":%><span><input name="res_<%=k%>" type="checkbox" value="<%=arr(j)%>" /><%=arr(j)%></span><%k=k+1%>
          <%case "text":%><span class="text"><%=arr(j)%><input name="res_<%=k%>" type="text" /></span><%k=k+1%>
        <%end select%>
      <%next%>
      </div>
      <%if arr(1)="radio" then%><%k=k+1%><%end if%>
    <%next%>
    <div class="bt"><input class="button" type="submit" value="<%=L_SUBMIT%>" /></div>
    </form>
  <%end if%>
  </div>
</div>
<!-- 新秀 -->
    </div>
    <div id="right">
      
<div class="here">
  <span><%=channel_title%></span>
  <div class="link">
  <a href="<%=S_ROOT%>"><%=L_HOME%></a>
  <%if channel_title<>"" then%><img src="<%=S_TPL_PATH%>images/here_img.gif" /><a href="<%=S_ROOT%>?<%=url_page%>.<%=suffix%>"><%=channel_title%></a><%end if%>
  <%if cat_name<>"" then%><img src="<%=S_TPL_PATH%>images/here_img.gif" /><a href="<%=S_ROOT%>?<%=url_page%>_<%=cat_id%>_0_0.<%=suffix%>"><%=cat_name%></a><%end if%>
  <%if page_title<>"" then%><img src="<%=S_TPL_PATH%>images/here_img.gif" /><a href="<%=S_ROOT%>?<%=url%>"><%=left(page_title,20)%></a><%end if%>
  </div>
</div>
<!-- 新秀 -->
      
<div class="mes_sheet">
  <ul>
    <%for i=0 to message_u%>
    <li>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="30"><span><%=message(i,0)%></span>&nbsp;<%=message(i,1)%>&nbsp;<%=L_SAID%>：</td>
        </tr>
        <tr>
          <td><div><%=message(i,2)%></div></td>
        </tr>
        <%if message(i,3)<>"" then%>
        <tr>
          <td><div class="reply"><span><%=L_REPLY%>：</span><%=message(i,3)%></div></td>
        </tr>
        <%end if%>
      </table>
    </li>
    <%next%>
  </ul>
  <%if message_u<0 then%><div class="not_found"><%=L_NOT_FOUND%></div><%end if%>
</div>
<div class="page_link"><div class="in">
  <span><%=L_ALL%><%=page_sum%><%=L_PAGES%></span>
  <span><%=L_NO%><%=page%><%=L_PAGE%></span>
  <%if page_sum<>1 then%>
    <a href="<%=S_ROOT%>?message_<%=cat%>_1_0.<%=suffix%>"><%=L_FIRST_PAGE%></a>
    <%if page-1>0 then%><a href="<%=S_ROOT%>?message_<%=cat%>_<%=page-1%>_0.<%=suffix%>"><%=L_PREVIOUS_PAGE%></a><%end if%>
    <%if page-4>0 then%><a class="num" href="<%=S_ROOT%>?message_<%=cat%>_<%=page-4%>_0.<%=suffix%>">【<%=page-4%>】</a><%end if%>
    <%if page-3>0 then%><a class="num" href="<%=S_ROOT%>?message_<%=cat%>_<%=page-3%>_0.<%=suffix%>">【<%=page-3%>】</a><%end if%>
    <%if page-2>0 then%><a class="num" href="<%=S_ROOT%>?message_<%=cat%>_<%=page-2%>_0.<%=suffix%>">【<%=page-2%>】</a><%end if%>
    <%if page-1>0 then%><a class="num" href="<%=S_ROOT%>?message_<%=cat%>_<%=page-1%>_0.<%=suffix%>">【<%=page-1%>】</a><%end if%>
    <a class="num" href="<%=S_ROOT%>?message_<%=cat%>_<%=page%>_0.<%=suffix%>" style="color:#7fb402;">【<%=page%>】</a>
    <%if page+1<=page_sum then%><a class="num" href="<%=S_ROOT%>?message_<%=cat%>_<%=page+1%>_0.<%=suffix%>">【<%=page+1%>】</a><%end if%>
    <%if page+2<=page_sum then%><a class="num" href="<%=S_ROOT%>?message_<%=cat%>_<%=page+2%>_0.<%=suffix%>">【<%=page+2%>】</a><%end if%>
    <%if page+3<=page_sum then%><a class="num" href="<%=S_ROOT%>?message_<%=cat%>_<%=page+3%>_0.<%=suffix%>">【<%=page+3%>】</a><%end if%>
    <%if page+4<=page_sum then%><a class="num" href="<%=S_ROOT%>?message_<%=cat%>_<%=page+4%>_0.<%=suffix%>">【<%=page+4%>】</a><%end if%>
    <%if page+1<=page_sum then%><a href="<%=S_ROOT%>?message_<%=cat%>_<%=page+1%>_0.<%=suffix%>"><%=L_NEXT_PAGE%></a><%end if%>
    <a href="<%=S_ROOT%>?message_<%=cat%>_<%=page_sum%>_0.<%=suffix%>"><%=L_LAST_PAGE%></a>
  <%end if%>
  <form name="form_jump" action="" method="get">
    <input class="text" type="text" name="page" value="" />
    <input class="button" type="button" onclick="page_jump()" value="<%=L_JUMP%>"/>
  </form>
  <script language="javascript">
  function page_jump()
  {
	  var val = document.form_jump.page.value;
	  document.location.href = "<%=S_ROOT%>?message_<%=cat%>_" + val + "_0.<%=suffix%>";
  }
  </script>
</div></div>
<div id="leave_word">
  <form action="<%=S_ROOT%>index/submit.asp" method="post">
    <input name="cmd" type="hidden" value="leave_word"/>
    <div>
      <input name="user" class="text" type="text" maxlength="30" value="<%=L_TRAVELLER%>"/>
      <input name="show" type="checkbox" value="2"/><%=L_SECRET%>
    </div>
    <div>
      <textarea name="text" cols="60" rows="5"></textarea>
    </div>
    <div>
      <input class="button" name="submit" type="submit" value="<%=L_SUBMIT%>" />
      <input class="button" name="reset" type="reset" value="<%=L_RESET%>"/>
    </div>
  </form>
</div>
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
