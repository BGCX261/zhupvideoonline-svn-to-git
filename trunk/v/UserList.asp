<!--#include file="Inc/syscode.asp"-->
<!-- #include file="Inc/Head.asp" -->
<div id="body">
  <table class="bodytable">
    <tr>
      <td class="bodyleft"><div id="leftbg">
          <!-- #include file="L_login.asp" -->
        </div>
        <img src="img/1.gif" /></td>
      <td class="bodyline"></td>
      <td class="bodyright"><table class="mbox">
          <tr>
            <td><div class="rtitle"><img src="img/title_ico.gif" /> �� �� �� �� </div></td>
          </tr>
          <tr>
            <td><b>�û�����</b></td>
          </tr>
          <tr>
            <td>
              <form method="get" name="myform" action="UserList.asp">
                <label for="uName">�û���</label><input name="uName" type="text" size="12" />
                <label for="cName">��ʵ����</label><input name="cName" type="text" size="12" />
                <label for="uStatus">���״̬</label>
                <select name="uStatus">
                  <option selected="selected" value="ȫ��">ȫ��</option>
                  <option value="True">ͨ��</option>
                  <option value="False">δͨ��</option>
                </select>
                <input name="input" type="submit" value=" �� ѯ " />
              </form>
            </td>
          </tr>
          <tr>
            <td id="newslist" class="cttd"><%
flag="��δ����"
set rs=server.createobject("adodb.recordset")


uName=Trim(request("uName"))
cName=Request("cName")
uStatus=Request("uStatus")
	sqltext="select * from User where UserID>0"
if uName="" then
	sqltext=sqltext
else
	sqltext=sqltext & " and UserName like '%" & uName & "%' "
end if

if cName="" then
	sqltext=sqltext
else
	sqltext=sqltext & " and Comane like '%" & cName & "%' "
end if

if uStatus="" or uStatus="ȫ��" then
	sqltext=sqltext
else
	sqltext=sqltext & " and LockUser = "  & uStatus
end if
sqltext=sqltext & " order by UserID desc"


rs.open sqltext,conn,1,1
dim PerPage
PerPage=20
'����û������ʱ
If rs.eof and rs.bof then
call showpages
response.write "<p align='center'><font color='#ff0000'>��û�κ�����</font></p>"


response.end
End if
'ȡ��ҳ��,���ж��û�������Ƿ��������͵����ݣ��粻�ǽ��Ե�һҳ��ʾ
text="0123456789"
Rs.PageSize=PerPage
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

If Rs.recordcount > PerPage then
call showpages
end if

'��ʾ���ӵ��ӳ���
Sub list()
%>
              <table width="100%">
              <tr>
                <td colspan="5"><b>�û��б�</b></td>
              </tr>
              <tr>
                <td>�û���</td>
                <td>��ʵ����</td>
                <td>רҵ�༶</td>
                <td>ע��ʱ��</td>
                <td>״̬</td>
              </tr>
                <%
if not rs.eof then
do while not rs.eof
%>
                <tr><td><a href="Server.asp"><%=rs("UserName")%></a></td><td><%=rs("Comane")%></td><td><%=rs("Add")%></td><td><font color="#0066CC"><%=rs("RegDate")%></font>&nbsp;
                  <%if formatdatetime(rs("RegDate"),2) = date() then%>
                  <strong><font color="#FF0000" face="Arial">New</font></strong>
                  <%end if%></td>
                  <td><%if rs("LockUser")="False" then Response.write "<font color='#ff0000'>���ͨ��</font>" else Response.write "<font color='#0000ff'>δ��˻�����</font>"%></td>
                </tr>
                <%
i=i+1
if i >= Perpage then exit do
rs.movenext
loop
end if
%>
              </table>			</td>
          </tr>  
          <tr>
		    <td id="tpage">
              <img src="img/arrow.gif" />
                <%
Response.write "��" & "<span id='numnews'>" & Cstr(Rs.RecordCount) & "</span>" & "ƪ"
Response.write "��" & "<span id='numnews'>" & Cstr(CurrentPage) &  "</span>" & "/" & Cstr(rs.pagecount) 
If currentpage > 1 Then
response.write "<span id='light'><a href='UserList.asp?&page="+cstr(1)+"'>��ҳ</a></span>"
Response.write "<span id='light'><a href='UserList.asp?page="+Cstr(currentpage-1)+"'>��һҳ</a></span>"
Else
Response.write "<span id='gray'>��һҳ </span>"
End if
If currentpage < Rs.PageCount Then
Response.write "<span id='light'><a href='UserList.asp?page="+Cstr(currentPage+1)+"'>��һҳ</a></span>"
Response.write "<span id='light'><a href='UserList.asp?page="+Cstr(Rs.PageCount)+"'>βҳ</a></span>"
Else
Response.write "<span id='gray'>��һҳ</span>"
End if
'response.write "<span id='goto' >ת��:<input type='text' name='page' size=4 maxlength=4 class=smallInput value="&Currentpage&">"
'response.write "<input class=buttonface type='submit'  value='Go'  name='cndok'></span>"
%>
              <%
End sub
rs.close
%>            </td>
          </tr>
        </table></td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
