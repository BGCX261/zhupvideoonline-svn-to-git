<!--#include file="bsconfig.asp"-->

<%
method=Request.QueryString("method")
pageno=Request.QueryString("pageno")
orderby=Request.QueryString("orderby")
select case orderby
	case "zs"
		ordertxt=" order by zs desc" 
	case else
		ordertxt=" order by dday desc" 
end select

htmltitle="��վ�������"
htmlname="sysadmin_count_report.asp?flag=1"
set rs=server.CreateObject("ADODB.Recordset")

sqltext="select querycount from syswork where code='101'"
rs.Open sqltext,conn,1,1
totcount=rs(0)
rs.Close
%>
<html>
<head>
<title><%=htmltitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body {font-size: 12px; color: #000; font-family: ����}
td {font-size: 12px; color: #000; font-family: ����;}

.t1 {font:12px ����;color=000000} 
.t2 {font:12px ����;color:ffffff} 
.bt1 {font:14px ����;color=000000} 
.bt2 {font:14px ����;color:ffffff} 

A.r1:link {font-size:12px;text-decoration:underline;color:#000000;}
A.r1:visited {font-size:12px;text-decoration:underline;color:#000000;}
A.r1:hover {font-size:12px;text-decoration:underline;color:#cc0000;}

A.r2:link {font-size:12px;text-decoration:underline;color:#ffffff;}
A.r2:visited {font-size:12px;text-decoration:underline;color:#ffffff;}
A.r2:hover {font-size:12px;text-decoration:underline;color:#ff6600;}

A.r3:link {font-size:12px;text-decoration:none;color:#000000;}
A.r3:visited {font-size:12px;text-decoration:none;color:#000000;}
A.r3:hover {font-size:12px;text-decoration:none;color:#cc0000;}

A.r4:link {font-size:12px;text-decoration:none;color:#ffffff;}
A.r4:visited {font-size:12px;text-decoration:none;color:#ffffff;}
A.r4:hover {font-size:12px;text-decoration:underline;color:#cc0000;}

A.r5:link {font-size:12px;text-decoration:underline;color:#000077;}
A.r5:visited {font-size:12px;text-decoration:underline;color:#000077;}
A.r5:hover {font-size:12px;text-decoration:underline;color:#ff0000;}

A.r6:link {font-size:12px;text-decoration:none;color:#ff6600;}
A.r6:visited {font-size:12px;text-decoration:none;color:#ff6600;}
A.r6:hover {font-size:12px;text-decoration:underline;color:#cc0000;}

-->
</style>

</head>
<body topmargin=5>
<table width="100%" align=center>
	<tr>
	<td valign="top" align="left" width="100%">
		<form name="form1">
		<table width="100%" height="20" border="0">
			<tr>
			<td style="font-size:12px;"></td>
			<td><font color="navy" style="font-size:14px"><b><%=htmltitle%></b>(����<font color=red><%=totcount%></font>�˴η���)</font></td>
			<td align="right">
			</td>
			</tr>
		</table>
		<table width="100%" height="20" border="0">
			<tr>
		<%
		sqltext="select count(*) as zs,convert(char(10),logindate,20) as dday from webcount group by convert(char(10),logindate,20) "+ordertxt
		'Response.Write sqltext
		rs.Open sqltext,conn,1,1

		listcs=30
		rs.PageSize=listcs
		if pageno="" then 
			pageno=1
		else
			pageno=cint(pageno)	
		end if	
		if rs.PageCount>0 then rs.AbsolutePage=pageno
	%>
		<td class=t1 align="left" width=200>
			<a class=r5 href="Javascript:window.location.reload()">[ˢ���б�]</a>
			<a class=r5 href="sysadmin_count_list.asp">[�����б�]</a>
			</td>
		<td  align="right">
		<%if rs.PageCount>1 then%>
			[��<b><%=pageno%></b>ҳ, ��<b><%=rs.PageCount%></b>ҳ<b><%=rs.RecordCount%></b>����¼]		              
			<%if pageno>1 then%>
				<a class=r1 href="<%=htmlname%>&pageno=1">��ҳ</a>              
				<a class=r1 href="<%=htmlname%>&pageno=<%=pageno-1%>">��ҳ</a>              
			<%end if              
			if pageno<rs.PageCount then%>              
				<a class=r1 href="<%=htmlname%>&pageno=<%=pageno+1%>">��ҳ</a>              
				<a class=r1 href="<%=htmlname%>&pageno=<%=rs.PageCount%>">ĩҳ</a>              
			<%end if%>              
		<%end if              
		%>              
		</td>
		</tr>
		</table>
			<table width="100%" border="0" cellspacing="1" cellpadding="1" style="font-size:14.5px;line-height:100%">
				<tr bgcolor="336699" style="color:#ffffff" align="center">
				<td class=bt2><a class=r2 href="<%=htmlname%>&orderby=rq">����</a></td><td class=bt2><a class=r2 href="<%=htmlname%>&orderby=zs">���ʴ���</a></td><td></td>
				</tr>
		<%
		i=0
		do while not rs.EOF and i<listcs 
			if (i mod 2)=1 then
				Response.Write("<tr bgcolor=fefefe>")
			else
				Response.Write("<tr bgcolor=efefef>")
			end if	
		%>
			<td width="180" align=right><%=rs(1)%>(<%=Weekdayname(Weekday(cdate(rs("dday"))),true)%>)</td>
			<td width="120" align=right><%=rs(0)%></td>
			<td></td>
			</tr>
		<%	
			i=i+1
			rs.MoveNext
		loop
		%>
			</table>
		</td></tr>
	<tr><td>
		<table width="100%" cellspacing="0" cellpadding="0" border=0> 
		<tr><td><hr size=1></td></tr>
			<tr><td align=right>
			[��<b><%=pageno%></b>ҳ, ��<b><%=rs.PageCount%></b>ҳ<b><%=rs.RecordCount%></b>����¼]			
			<%if pageno>1 then%>
				<a class=r1 href="<%=htmlname%>&pageno=1">[��ҳ]</a>              
				<a class=r1 href="<%=htmlname%>&pageno=<%=pageno-1%>">[��ҳ]</a>              
			<%end if              
			if pageno<rs.PageCount then%>              
				<a class=r1 href="<%=htmlname%>&pageno=<%=pageno+1%>">[��ҳ]</a>              
				<a class=r1 href="<%=htmlname%>&pageno=<%=rs.PageCount%>">[ĩҳ]</a>              
			<%end if%>              
			</td></tr>              
		</table>
	</td></tr>
		</form>
</td></tr><table>
</body>
</html>

<script language="Javascript">
	function delrec()
	{
		if (confirm('��ȷ��Ҫ�����վ���ʼ�¼��')==true)
		{
		window.location.href="<%=htmlname%>&method=clearall"
		return false;
		}
	}

</script>

