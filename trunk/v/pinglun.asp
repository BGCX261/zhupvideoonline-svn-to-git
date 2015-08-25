<!--#include file="Inc/syscode.asp"-->
<!--#include file="Inc/top.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%
pages=10		'每页评论数
ArticleID=trim(request("ArticleID"))
%>


<%
'set rs=conn.execute("select * from Product where ArticleID='"&ArticleID&"'")
'if not (rs.eof and rs.bof) then response.write "以下是【"&rs("Title")&"】的所有用户评论：<br><br>"

response.write ArticleID
sqlClassWrd="select * from Product where ArticleID=clng(ArticleID) "
Set rsClassWrd= Server.CreateObject("ADODB.Recordset")
rsClassWrd.open sqlClassWrd,conn,1,1
if not (rsClassWrd.eof and rsClassWrd.bof) then response.write "以下是【"&rsClassWrd("Title")&"】的所有用户评论：<br><br>"


set rs=Server.CreateObject("ADODB.RecordSet")
sql="select * from buyok_pinglun where prodid='"&ArticleID&"' order by AddDate desc"
rs.open sql,conn,1,3
if rs.eof and rs.bof then
response.write "<br>暂时没有用户对此作品提交评论。<br><br>"
else
rs.pageSize = pages '每页记录数
allPages = rs.pageCount	'总页数
page = Request("page")	'从浏览器取得当前页

'if是基本的出错处理

If not isNumeric(page) then page=1

if isEmpty(page) or clng(page) < 1 then
page = 1
elseif clng(page) >= allPages then
page = allPages 
end if
rs.AbsolutePage = page			'转到某页头部
do while not rs.eof and pages>0
%>

<table cellSpacing=3 cellPadding=3 width=100% align=center bgColor=#F7FAFD border=0 onMouseOver="bgColor='#ebebeb'" onMouseOut="bgColor='#F7FAFD'">
<%
response.write "<tr><td width=33% >姓名："&rs("name")&"<td widht=33% ><a href='mailto:"&rs("mail")&"'><img border=0 src=images/small/e-mail.gif></a></td><td width=33% align=right>"&ii11ii1(rs("IP"))&"&nbsp;&nbsp;&nbsp;</td></tr>"
response.write "<tr><td>时间："&rs("adddate")&"</td><td> </td><td align=right>"
for j=1 to rs("jibie")
response.write "<font color=orange>☆</font>"
next
	nr=rs("nr")			'内容
	'bad1=split(bad,"/")		'过滤脏话
	'for t=0 to ubound(bad1)
	'nr=replace(nr,bad1(t),"***")
	'next
response.write "&nbsp;&nbsp;&nbsp;</td></tr><tr><td colspan=3><font color=darkgray><span style='line-height=100%'>"&replace(server.htmlencode(nr),vbCRLF,"<BR>")&"</span></font></td></tr></table>"
pages = pages - 1
rs.movenext
if rs.eof then exit do
loop
end if

response.write "<br>总计"&RS.RecordCount&"条评论 "
if page = 1 then
response.write "<font color=darkgray>首页 前页</font>"
else
response.write "<a href=pinglun.asp?ArticleID="&prodid&"&page=1>首页</a> <a href=pinglun.asp?ArticleID="&prodid&"&page="&page-1&">前页</a>"
end if
if page = allpages then
response.write "<font color=darkgray> 下页 末页</font>"
else
response.write " <a href=pinglun.asp?ArticleID="&prodid&"&page="&page+1&">下页</a> <a href=pinglun.asp?ArticleID="&prodid&"&page="&allpages&">末页</a>"
end if
response.write " 第"&page&"页 共"&allpages&"页 "

conn.close
set conn=nothing
%>
</td></tr>
</table>
