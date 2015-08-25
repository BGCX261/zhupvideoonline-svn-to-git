<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>∷<%=BsCompanyName%>∷</title>
<link href="inc/cnstyle.css" type="text/css" rel="stylesheet" />
<link href="inc/adcss.css" rel="stylesheet" type="text/css">
</head>
<body>
<div class="top">
    <div id="headnav">
        <div class="hTop"><a href="http://www.zhcpt.edu.cn/" target="_blank">学院主页</a> | <a onClick="window.external.addFavorite('http://<%=BsWeb%>/','<%=BsCompanyName%>');return false;" href="http://<%=BsWeb%>/">收藏本站</a> |
              <!--#include file="gb2big5.htm"-->
        </div>
        <div class="logobg"><img src="images/logo.png"></div>
    <!--菜单开始-->
        <%
        menunav=replace(rs_BsCo("menunav"),"../","./")
        response.write menunav
        %>
    <!--菜单结束-->    
    </div>
    
    <div id="header">
            <script src="js/yu.js" type="text/javascript"></script>
            <script src="js/tb.js" type="text/javascript"></script>
            <div id="MainPromotionBanner"  title="<%=BsCompanyName%>">
                <div id="SlidePlayer">
                    <ul class="Slides">
                        <li><img src="images/headbg1.jpg"></li>
                        <li><img src="images/headbg2.jpg"></li>
                        <li><img src="images/headbg3.jpg"></li>
                    </ul>
                </div>
                <script type="text/javascript">
                        TB.widget.SimpleSlide.decoration('SlidePlayer', {eventType:'mouse', effect:'scroll'});
                 </script>
            </div>
    </div>
</div>

