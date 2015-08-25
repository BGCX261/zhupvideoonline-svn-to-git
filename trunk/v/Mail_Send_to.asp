<!--#include file="Inc/syscode.asp" -->
<!-- #include file="Inc/Head.asp" -->
<table align="center" class="a2">
	<tr>
		<td class="a1">发 送 邮 件</td>
	</tr>
	<tr class="a4">
		<td>
	  <% 
'发送
set rs=server.createobject("adodb.recordset") 
sql="select * from email "
rs.open sql,conn,1,3  

'读取默认的邮件标题及内容 
set rs1=server.createobject("adodb.recordset")
sql1="select * from maildefault "
rs1.open sql1,conn,1,3  

'检验发信人地址
tomail=request("tomail")
if tomail="" then
Response.Write("<script language=""JavaScript"">alert(""错误：您没输入邮件地址，请返回检查！！"");history.go(-1);</script>")
response.end
end if

'检验邮件主题
mailsubject=request("mailsubject")
if mailsubject="" then
Response.Write("<script language=""JavaScript"">alert(""错误：您没输入邮件主题，请返回检查！！"");history.go(-1);</script>")
response.end
end if

'设置邮件内容
mailbody=request("mailbody")
if mailbody="" then
Response.Write("<script language=""JavaScript"">alert(""错误：您没输入邮件内容，请返回检查！！"");history.go(-1);</script>")
response.end
end if

'判断对谁发信
frommail=request("frommail")
'写发信信息
response.write "发信人地址: "&tomail
response.write "<br><br><br>"
if frommail<>"" then
response.write "收信人地址："&frommail
else
response.write "正在进行邮件发送！"
end if
if frommail="" then
'对于设定邮箱发信
Response.Write("<script language=""JavaScript"">alert(""错误：发送错误，请返回检查！！"");history.go(-1);</script>")
response.end
else
Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
objCDOMail.From = tomail
objCDOMail.To = frommail
objCDOMail.Subject = mailsubject  
objCDOMail.Body = mailbody   
objCDOMail.Send
Set objCDOMail = Nothing
end if
response.write "<br><br><br>"
response.write "邮件发送成功！^&^"
'response.write "<br><br><br>"
'response.write rs1("mailsubject")
%> 
		</td>
	</tr>
	<tr>
	  <td class="iptsubmit"><a href="Mail_Send.asp">返回</a></td>
	</tr>
</table>
<%htmlend%>
