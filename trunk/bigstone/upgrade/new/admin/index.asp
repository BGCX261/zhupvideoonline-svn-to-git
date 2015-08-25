<!-- #include file="../include/common.asp" -->
<!-- #include file="system/lang.asp" -->
<%if get_session("user","") <> "" and get_session("password","") <> "" then response.redirect("system/main.asp")%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<title><%=L_PAGE_TITLE%></title>
<link href="system/css.css" rel="stylesheet" type="text/css">
</head>
<body id="login">
<div class="box">
  <div>
  <form method="post" action="system/main.asp">
    <table cellspacing="0" cellpadding="0">
      <tr>
        <td>用户名：</td>
        <td><input name="user" type="text" class="text" maxlength="30" size="15"></td>
        <td width="70px"></td>
      </tr>
      <tr>
        <td>密&nbsp;&nbsp;码：</td>
        <td><input name="password" type="password" class="text" size="15"></td>
        <td><input name="submit" type="submit" value="登录"></td>
      </tr>
    </table>
  </form>
  </div>
</div>
</body>
</html>
<%'新秀%>