<%if get_session("login","") = "" or S_ROOT = "" then response.end%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<title><%=L_PAGE_TITLE%></title>
<link href="../system/css.css" rel="stylesheet" type="text/css">
<style type="text/css">html,body{overflow-x:hidden;overflow-y:scroll;}</style>
<script type="text/javascript" src="../system/script.js"></script>
</head>
<body id="main">
<!-- #include file="box.asp" -->

<%'新秀%>