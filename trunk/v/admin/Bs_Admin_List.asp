<!--#include file="bsconfig.asp"-->

<%if Request.QueryString("no")="eshop" then

dim SQL, Rs, contentID,CurrentPage
CurrentPage = request("Page")
contentID=request("ID")

set rs=server.createobject("adodb.recordset")
sqltext="delete from Bs_User where ID="& contentID
rs.open sqltext,conn,3,3
set rs=nothing
conn.close
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('�ɹ�ɾ����');" & Chr(13)		
		response.write "</script>" & Chr(13)
		response.redirect "Bs_Admin_List.asp"
end if

%>
<%

set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_User order by ID desc"
rs.open sqltext,conn,1,1

dim MaxPerPage
MaxPerPage=20

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
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("04")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">�� ̨ �� �� Ա �� ��</td>
	</tr>
	<tr class="a4">
		<td>
              <table width="100%">
                <tr class="opttitle"> 
                  <td> ID</td>
                  <td> �û���</td>
                  <td> ��ʵ����</td>
                  <td> ��ϵ�绰</td>
                  <td> �û����</td>
                  <td> ����Ժϵ��</td>
                  <td> רҵ�༶�������</td>
                  <td> ��ѧ���</td>
                  <td> ����</td>
                </tr>
                <%
if not rs.eof then
i=0
do while not rs.eof
%>
                <tr class="opt"> 
                  <td> <%=rs("ID")%></td>
                  <td><%=rs("adminName")%></td>
                  <td><%=rs("realname")%></td>
                  <td><%=rs("uphone")%></td>
                  <td><%=rs("adminRole")%></td>
                  <td><%=rs("ufaculty")%></td>
                  <td><%=rs("ubelong")%></td>
                  <td><%=rs("admission")%></td>
                  <td><a href="Bs_Admin_edit.asp?id=<%=rs("ID")%>">�޸�</a> | <a href="Bs_Admin_list.asp?id=<%=rs("ID")%>&amp;no=eshop">ɾ��</a></td>
                </tr>
                <% 
i=i+1 
if i >= MaxPerpage then exit do 
rs.movenext 
loop 
end if 
%>
                <tr> 
                  <td colspan="9" id="turnpage"><%
Response.write "<strong><font color='#000000'>-> ȫ��-</font>"
Response.write "��</font>" & "<font color=#FF0000>" & Cstr(Rs.RecordCount) & "</font>" & "<font color='#000000'>��</font></strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.write "<strong><font color='#000000'>��</font>" & "<font color=#FF0000>" & Cstr(CurrentPage) &  "</font>" & "<font color='#000000'>/" & Cstr(rs.pagecount) & "</font></strong>&nbsp;"
If currentpage > 1 Then
response.write "<strong><a href='Bs_Admin_list.asp?&page="+cstr(1)+"'><font color='#000000'>��ҳ</font></a><font color='#ffffff'> </font></strong>"
Response.write "<strong><a href='Bs_Admin_list.asp?page="+Cstr(currentpage-1)+"'><font color='#000000'>��һҳ</font></a><font color='#ffffff'> </font></strong>"
Else
Response.write "<strong><font color='#000000'>��һҳ </font></strong>"
End if
If currentpage < Rs.PageCount Then
Response.write "<strong><a href='Bs_Admin_list.asp?page="+Cstr(currentPage+1)+"'><font color='#000000'>��һҳ</font></a><font color='#ffffff'> </font>"
Response.write "<a href='Bs_Admin_list.asp?page="+Cstr(Rs.PageCount)+"'><font color='#000000'>βҳ</font></a></strong>&nbsp;&nbsp;"
Else
Response.write ""
Response.write "<strong><font color='#000000'>��һҳ</font></strong>&nbsp;&nbsp;"
End if
'response.write "</td><td align='right'>"
'response.write "<font color='#000000' >ת����</font><input type='text' name='page' size=4 maxlength=4 class=smallInput value="&Currentpage&">&nbsp;"
'response.write "<input class=buttonface type='submit'  value='Go'  name='cndok'>&nbsp;&nbsp;"
%></td>
                </tr>
              </table>
			  <%
End sub
rs.close

%>
		</td>
	</tr>
</table>
<%htmlend%>
