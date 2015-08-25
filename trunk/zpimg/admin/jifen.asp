<!--#include file="../include/head.asp"-->
<!--#include file="../include/ofunctions.asp"-->
<%If User_Group<>"admin" Then TurnTo "../"%>
<%
'** 文件说明：修改会员积分
Dim id,uid,UserName,Bonus
Act=Request.QueryString("act")
id=Request.QueryString("id")

If act="" And id="" Then
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改会员积分</title>
<link href="images/css.css" rel="stylesheet" type="text/css">
<body >
<form name="form1" method="POST" action="jifen.asp?act=EditBonus">
<div class="JiFen">
<p class="P1">用户名：<input name="id" type="text" value="" class="IT2"></p>
<p class="P1">新积分：<input name="Bonus" type="text" value="" class="IT2"><input type="submit" value="修改" class="btn2"></p>
<p class="G2"></p>
<p class="P1 red">说明：此方法为直接修改积分，并非加减积分</p>
</div>
</form>
</body>
</html>
<%
End If

If act="" And id<>"" Then
Sql="select * from Users where UserName='"&id&"'"
OpenRs(Sql)
Bonus=Rs("Bonus")
uid=Rs("id")
Rs.Close
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改会员积分</title>
<link href="images/css.css" rel="stylesheet" type="text/css">
<body >
<form name="form1" method="POST" action="jifen.asp?act=EditBonus">
<div class="JiFen">
<p class="P1">用户名：<%=id%><input type="hidden" name="id" value="<%=id%>"></p>
<p class="P1">积分数：<%=Bonus%></p>
<p class="P1">新积分：<input name="Bonus" type="text" value="" class="IT2"><input type="submit" value="修改" class="btn2"></p>
<p class="G2"></p>
<p class="P1 red">说明：此方法为直接修改积分，并非加减积分</p>
</div>
</form>
</body>
</html>
<%
End If

If act="EditBonus" Then
Bonus=Trim(Request.Form("Bonus"))
id=Trim(Request.Form("id"))
If Bonus="" Then Call ShowError("请输入积分！",1)

Sql="select * from Users where UserName='"&id&"'"
OpenRs(Sql)
Rs("Bonus")=Bonus
rs.update
rs.close

'创建管理员操作日志
Sql="select * from SetLog"
OpenRs(Sql)
Rs.AddNew
Rs("AdminName")=User_ID
Rs("SetTxt")="修改用户["&id&"]的积分"
Rs("SetTime")=Now
Rs.Update
Rs.Close

Response.write("<script language='javascript'>alert('修改成功！');window.location='jifen.asp?id="&id&"';</script>")

End If
CloseAll
%>