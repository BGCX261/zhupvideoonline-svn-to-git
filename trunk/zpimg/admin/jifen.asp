<!--#include file="../include/head.asp"-->
<!--#include file="../include/ofunctions.asp"-->
<%If User_Group<>"admin" Then TurnTo "../"%>
<%
'** �ļ�˵�����޸Ļ�Ա����
Dim id,uid,UserName,Bonus
Act=Request.QueryString("act")
id=Request.QueryString("id")

If act="" And id="" Then
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�޸Ļ�Ա����</title>
<link href="images/css.css" rel="stylesheet" type="text/css">
<body >
<form name="form1" method="POST" action="jifen.asp?act=EditBonus">
<div class="JiFen">
<p class="P1">�û�����<input name="id" type="text" value="" class="IT2"></p>
<p class="P1">�»��֣�<input name="Bonus" type="text" value="" class="IT2"><input type="submit" value="�޸�" class="btn2"></p>
<p class="G2"></p>
<p class="P1 red">˵�����˷���Ϊֱ���޸Ļ��֣����ǼӼ�����</p>
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
<title>�޸Ļ�Ա����</title>
<link href="images/css.css" rel="stylesheet" type="text/css">
<body >
<form name="form1" method="POST" action="jifen.asp?act=EditBonus">
<div class="JiFen">
<p class="P1">�û�����<%=id%><input type="hidden" name="id" value="<%=id%>"></p>
<p class="P1">��������<%=Bonus%></p>
<p class="P1">�»��֣�<input name="Bonus" type="text" value="" class="IT2"><input type="submit" value="�޸�" class="btn2"></p>
<p class="G2"></p>
<p class="P1 red">˵�����˷���Ϊֱ���޸Ļ��֣����ǼӼ�����</p>
</div>
</form>
</body>
</html>
<%
End If

If act="EditBonus" Then
Bonus=Trim(Request.Form("Bonus"))
id=Trim(Request.Form("id"))
If Bonus="" Then Call ShowError("��������֣�",1)

Sql="select * from Users where UserName='"&id&"'"
OpenRs(Sql)
Rs("Bonus")=Bonus
rs.update
rs.close

'��������Ա������־
Sql="select * from SetLog"
OpenRs(Sql)
Rs.AddNew
Rs("AdminName")=User_ID
Rs("SetTxt")="�޸��û�["&id&"]�Ļ���"
Rs("SetTime")=Now
Rs.Update
Rs.Close

Response.write("<script language='javascript'>alert('�޸ĳɹ���');window.location='jifen.asp?id="&id&"';</script>")

End If
CloseAll
%>