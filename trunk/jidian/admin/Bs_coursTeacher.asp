<!--#include file="bsconfig.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("19")%>
<%if Request.QueryString("no")="eshop" then

dim SQL, Rs, contentID,CurrentPage
CurrentPage = request("Page")
contentID=request("ID")

set rs=server.createobject("adodb.recordset")

call checkmanage("22")

sqltext="delete from Teacher where id="& contentID
rs.open sqltext,conn,3,3
set rs=nothing
conn.close
response.redirect "Bs_coursTeacher.asp"
end if

%>
<%

set rs=server.createobject("adodb.recordset")
sqltext="select * from Teacher order by sortId desc"
rs.open sqltext,conn,1,1

dim MaxPerPage
MaxPerPage=10

dim text,checkpage
text="0123456789"
Rs.PageSize=MaxPerPage
for i=1 to len(request("page"))
checkpage=instr(1,text,mid(request("page"),i,1))
if checkpage=0 then
exit for
end if
next

If checkpage<>0 then
If NOT IsEmpty(request("page")) Then
CurrentPage=Cint(request("page"))
If CurrentPage < 1 Then CurrentPage = 1
If CurrentPage > Rs.PageCount Then CurrentPage = Rs.PageCount
Else
CurrentPage= 1
End If
If not Rs.eof Then Rs.AbsolutePage = CurrentPage end if
Else
CurrentPage=1
End if

call showpages
call list

If Rs.recordcount > MaxPerPage then
call showpages
end if

'��ʾ���ӵ��ӳ���
Sub list()%>

<BR>
<table cellpadding="2" cellspacing="1" border="0" width="560" align="center" class="a2">
	<tr>
		<td class="a1" height="25" align="center"><strong>��ʦ��Ϣ����</strong></td>
	</tr>
	<tr class="a4">
		<td>
						<table width="100%" border="0" cellspacing="2" cellpadding="3">
							<tr class="opttitle"> 
								<td width="5%"> ����ID</td>
								<td width="10%"> ��ʦ����</td>
                                <td width="35%"> �����γ�</td>
                                <td width="30%"> ��ʦ����</td>
								<td width="10%"> ��ʦͼƬ</td>
								<td width="5%"> ����</td>
                                <td width="5%"> ����</td>
							</tr>
<%
if not rs.eof then
i=0
do while not rs.eof
%>
							<tr class="opt"> 
								<td><%=rs("sortId")%></td>
								<td><%=rs("teachName")%></td>
                                <td><%=rs("teachSubj")%></td>
                                <td><%=rs("teachSay")%></td>
								<td> <img src="../<%=rs("honor")%>" width="45" height="60" title="<%=rs("teachInfo")%>" /></td>
								<td> <a href="Bs_coursTeacher_edit.asp?id=<%=rs("id")%>">�޸�</a>
								 <a href="Bs_coursTeacher.asp?id=<%=rs("id")%>&no=eshop">ɾ��</a></td>
                                <td><a href="paixu3act.asp?page=<%=currentPage%>&act=up&tempo=<%=i+1%>&id=<%=rs("id")%>">�ϡ� </a> <a href="paixu3act.asp?page=<%=currentPage%>&act=down&tempo=<%=i+1%>&id=<%=rs("id")%>">�¡� </a></td>
							</tr>
<% 
i=i+1 
if i >= MaxPerpage then exit do 
rs.movenext 
loop 
end if 
%>
							<tr> 
								<td id="turnpage" colspan="7">
<%
Response.write "<strong><font color='#000000'>-> ȫ��-</font>"
Response.write "��</font>" & "<font color=#FF0000>" & Cstr(Rs.RecordCount) & "</font>" & "<font color='#000000'>λ��ʦ</font></strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.write "<strong><font color='#000000'>��</font>" & "<font color=#FF0000>" & Cstr(CurrentPage) &  "</font>" & "<font color='#000000'>/" & Cstr(rs.pagecount) & "</font></strong>&nbsp;"
If currentpage > 1 Then
response.write "<strong><a href='Bs_coursTeacher.asp?&page="+cstr(1)+"'><font color='#000000'>��ҳ</font></a><font color='#ffffff'> </font></strong>"
Response.write "<strong><a href='Bs_coursTeacher.asp?page="+Cstr(currentpage-1)+"'><font color='#000000'>��һҳ</font></a><font color='#ffffff'> </font></strong>"
Else
Response.write "<strong><font color='#000000'>��һҳ </font></strong>"
End if
If currentpage < Rs.PageCount Then
Response.write "<strong><a href='Bs_coursTeacher.asp?page="+Cstr(currentPage+1)+"'><font color='#000000'>��һҳ</font></a><font color='#ffffff'> </font>"
Response.write "<a href='Bs_coursTeacher.asp?page="+Cstr(Rs.PageCount)+"'><font color='#000000'>βҳ</font></a></strong>&nbsp;&nbsp;"
Else
Response.write ""
Response.write "<strong><font color='#000000'>��һҳ</font></strong>&nbsp;&nbsp;"
End if
'response.write "</td><td align='right'>"
'response.write "ת����<input type='text' name='page' size=4 maxlength=4 class=smallInput value="&Currentpage&"> "
'response.write "<input class=buttonface type='submit'  value='Go'  name='cndok'>"
%>
								</td>
							</tr>
						</table>
<%
End sub
rs.close

%>
		</td>
	</tr>
</table>
<BR>
<%htmlend%>

