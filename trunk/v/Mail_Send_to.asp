<!--#include file="Inc/syscode.asp" -->
<!-- #include file="Inc/Head.asp" -->
<table align="center" class="a2">
	<tr>
		<td class="a1">�� �� �� ��</td>
	</tr>
	<tr class="a4">
		<td>
	  <% 
'����
set rs=server.createobject("adodb.recordset") 
sql="select * from email "
rs.open sql,conn,1,3  

'��ȡĬ�ϵ��ʼ����⼰���� 
set rs1=server.createobject("adodb.recordset")
sql1="select * from maildefault "
rs1.open sql1,conn,1,3  

'���鷢���˵�ַ
tomail=request("tomail")
if tomail="" then
Response.Write("<script language=""JavaScript"">alert(""������û�����ʼ���ַ���뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if

'�����ʼ�����
mailsubject=request("mailsubject")
if mailsubject="" then
Response.Write("<script language=""JavaScript"">alert(""������û�����ʼ����⣬�뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if

'�����ʼ�����
mailbody=request("mailbody")
if mailbody="" then
Response.Write("<script language=""JavaScript"">alert(""������û�����ʼ����ݣ��뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if

'�ж϶�˭����
frommail=request("frommail")
'д������Ϣ
response.write "�����˵�ַ: "&tomail
response.write "<br><br><br>"
if frommail<>"" then
response.write "�����˵�ַ��"&frommail
else
response.write "���ڽ����ʼ����ͣ�"
end if
if frommail="" then
'�����趨���䷢��
Response.Write("<script language=""JavaScript"">alert(""���󣺷��ʹ����뷵�ؼ�飡��"");history.go(-1);</script>")
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
response.write "�ʼ����ͳɹ���^&^"
'response.write "<br><br><br>"
'response.write rs1("mailsubject")
%> 
		</td>
	</tr>
	<tr>
	  <td class="iptsubmit"><a href="Mail_Send.asp">����</a></td>
	</tr>
</table>
<%htmlend%>
