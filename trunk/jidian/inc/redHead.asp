<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��<%=BsCompanyName%>��</title>
<link href="inc/redstyle.css" type="text/css" rel="stylesheet" />
<link href="inc/adcss.css" rel="stylesheet" type="text/css">
</head>
<body>
<div class="top">
    <div id="headnav">
        <div class="hTop"><a href="http://www.zhcpt.edu.cn/" target="_blank">ѧԺ��ҳ</a> | <a onClick="window.external.addFavorite('http://<%=BsWeb%>/','<%=BsCompanyName%>');return false;" href="http://<%=BsWeb%>/">�ղر�վ</a> |
              <!--#include file="gb2big5.htm"-->
        </div>
        <div class="logobg"><img src="images/logo.png"></div>
    <!--�˵���ʼ-->
        <%
        menunav=replace(rs_BsCo("menunav"),"../","./")
        response.write menunav
        %>
    <!--�˵�����-->    
    </div>
    
</div>

<div id="header">
    <img src="images/headRedBg.jpg">
</div>

