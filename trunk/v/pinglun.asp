<!--#include file="Inc/syscode.asp"-->
<!--#include file="Inc/top.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%
pages=10		'ÿҳ������
ArticleID=trim(request("ArticleID"))
%>


<%
'set rs=conn.execute("select * from Product where ArticleID='"&ArticleID&"'")
'if not (rs.eof and rs.bof) then response.write "�����ǡ�"&rs("Title")&"���������û����ۣ�<br><br>"

response.write ArticleID
sqlClassWrd="select * from Product where ArticleID=clng(ArticleID) "
Set rsClassWrd= Server.CreateObject("ADODB.Recordset")
rsClassWrd.open sqlClassWrd,conn,1,1
if not (rsClassWrd.eof and rsClassWrd.bof) then response.write "�����ǡ�"&rsClassWrd("Title")&"���������û����ۣ�<br><br>"


set rs=Server.CreateObject("ADODB.RecordSet")
sql="select * from buyok_pinglun where prodid='"&ArticleID&"' order by AddDate desc"
rs.open sql,conn,1,3
if rs.eof and rs.bof then
response.write "<br>��ʱû���û��Դ���Ʒ�ύ���ۡ�<br><br>"
else
rs.pageSize = pages 'ÿҳ��¼��
allPages = rs.pageCount	'��ҳ��
page = Request("page")	'�������ȡ�õ�ǰҳ

'if�ǻ����ĳ�����

If not isNumeric(page) then page=1

if isEmpty(page) or clng(page) < 1 then
page = 1
elseif clng(page) >= allPages then
page = allPages 
end if
rs.AbsolutePage = page			'ת��ĳҳͷ��
do while not rs.eof and pages>0
%>

<table cellSpacing=3 cellPadding=3 width=100% align=center bgColor=#F7FAFD border=0 onMouseOver="bgColor='#ebebeb'" onMouseOut="bgColor='#F7FAFD'">
<%
response.write "<tr><td width=33% >������"&rs("name")&"<td widht=33% ><a href='mailto:"&rs("mail")&"'><img border=0 src=images/small/e-mail.gif></a></td><td width=33% align=right>"&ii11ii1(rs("IP"))&"&nbsp;&nbsp;&nbsp;</td></tr>"
response.write "<tr><td>ʱ�䣺"&rs("adddate")&"</td><td> </td><td align=right>"
for j=1 to rs("jibie")
response.write "<font color=orange>��</font>"
next
	nr=rs("nr")			'����
	'bad1=split(bad,"/")		'�����໰
	'for t=0 to ubound(bad1)
	'nr=replace(nr,bad1(t),"***")
	'next
response.write "&nbsp;&nbsp;&nbsp;</td></tr><tr><td colspan=3><font color=darkgray><span style='line-height=100%'>"&replace(server.htmlencode(nr),vbCRLF,"<BR>")&"</span></font></td></tr></table>"
pages = pages - 1
rs.movenext
if rs.eof then exit do
loop
end if

response.write "<br>�ܼ�"&RS.RecordCount&"������ "
if page = 1 then
response.write "<font color=darkgray>��ҳ ǰҳ</font>"
else
response.write "<a href=pinglun.asp?ArticleID="&prodid&"&page=1>��ҳ</a> <a href=pinglun.asp?ArticleID="&prodid&"&page="&page-1&">ǰҳ</a>"
end if
if page = allpages then
response.write "<font color=darkgray> ��ҳ ĩҳ</font>"
else
response.write " <a href=pinglun.asp?ArticleID="&prodid&"&page="&page+1&">��ҳ</a> <a href=pinglun.asp?ArticleID="&prodid&"&page="&allpages&">ĩҳ</a>"
end if
response.write " ��"&page&"ҳ ��"&allpages&"ҳ "

conn.close
set conn=nothing
%>
</td></tr>
</table>
