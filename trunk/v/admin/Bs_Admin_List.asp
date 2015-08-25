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
		response.write "alert('成功删除！');" & Chr(13)		
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

'显示帖子的子程序
Sub list()%>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("04")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">后 台 管 理 员 管 理</td>
	</tr>
	<tr class="a4">
		<td>
              <table width="100%">
                <tr class="opttitle"> 
                  <td> ID</td>
                  <td> 用户名</td>
                  <td> 真实姓名</td>
                  <td> 联系电话</td>
                  <td> 用户身份</td>
                  <td> 所属院系部</td>
                  <td> 专业班级或教研室</td>
                  <td> 入学年份</td>
                  <td> 操作</td>
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
                  <td><a href="Bs_Admin_edit.asp?id=<%=rs("ID")%>">修改</a> | <a href="Bs_Admin_list.asp?id=<%=rs("ID")%>&amp;no=eshop">删除</a></td>
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
Response.write "<strong><font color='#000000'>-> 全部-</font>"
Response.write "共</font>" & "<font color=#FF0000>" & Cstr(Rs.RecordCount) & "</font>" & "<font color='#000000'>条</font></strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.write "<strong><font color='#000000'>第</font>" & "<font color=#FF0000>" & Cstr(CurrentPage) &  "</font>" & "<font color='#000000'>/" & Cstr(rs.pagecount) & "</font></strong>&nbsp;"
If currentpage > 1 Then
response.write "<strong><a href='Bs_Admin_list.asp?&page="+cstr(1)+"'><font color='#000000'>首页</font></a><font color='#ffffff'> </font></strong>"
Response.write "<strong><a href='Bs_Admin_list.asp?page="+Cstr(currentpage-1)+"'><font color='#000000'>上一页</font></a><font color='#ffffff'> </font></strong>"
Else
Response.write "<strong><font color='#000000'>上一页 </font></strong>"
End if
If currentpage < Rs.PageCount Then
Response.write "<strong><a href='Bs_Admin_list.asp?page="+Cstr(currentPage+1)+"'><font color='#000000'>下一页</font></a><font color='#ffffff'> </font>"
Response.write "<a href='Bs_Admin_list.asp?page="+Cstr(Rs.PageCount)+"'><font color='#000000'>尾页</font></a></strong>&nbsp;&nbsp;"
Else
Response.write ""
Response.write "<strong><font color='#000000'>下一页</font></strong>&nbsp;&nbsp;"
End if
'response.write "</td><td align='right'>"
'response.write "<font color='#000000' >转到：</font><input type='text' name='page' size=4 maxlength=4 class=smallInput value="&Currentpage&">&nbsp;"
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
