<!--#include file="bsconfig.asp"-->

<%
set rs=server.CreateObject("adodb.recordset")
rs.open "Select * From scType",conn,1,3

if Request.QueryString("no")="yes" then
nbtyped= Trim(Request("nbtyped"))
Brief= Trim(Request("Brief"))
if nbtyped<>"" then
sql="delete from scType where nbtyped='" & nbtyped & "'"
conn.Execute  sql
sql="delete from Scientific where nbtyped='" & nbtyped & "'"
conn.Execute sql
end if


Set rs= Nothing
Set Conn = Nothing
Response.Redirect "Bs_scType.asp"
end if



if Request.QueryString("no")="eshop" then
nbtyped=request.form("nbtyped")
Brief=request.form("Brief")

If nbtyped="" Then
response.write "SORRY <br>"
response.write "������!"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from scType "
rs.open sql,conn,1,3
rs.addnew
rs("nbtyped")=nbtyped
rs("Brief")=Brief

rs.update
Response.Redirect "Bs_scType.asp"
end if
%>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("17")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">��ѧ����������</td>
	</tr>
	<tr class="a4">
		<td>
          <form method="post" action="Bs_scType.asp?no=eshop">
			<table width="100%">
                <tr> 
                  <td class="iptlab"> �����������</td>
                  <td class="ipt"><input type="text" name="nbtyped" size="35" maxlength="40"></td>
                </tr>
                <tr> 
                  <td class="iptlab"> ����˵��</td>
                  <td class="ipt"><textarea name="Brief" cols="50" rows="4"></textarea></td>
                </tr>
                <tr> 
                  <td colspan="2" class="iptsubmit">  
                      <input type="submit" name="Submit" value="ȷ��">
                      <input type="reset" name="Submit2" value="����">
                    </td>
                </tr>
              </table>
              <br> 
            <table width="100%">
                <tr class="opttitle"> 
                  <td>�������</td>
                  <td>����</td>
                  <td>����</td>
                </tr>
                <% do while not rs.eof %>
                <tr class="opt"> 
                  <td><%=rs("nbtyped")%></td>
                  <td><a href="Bs_scType_edit.asp?id=<%=rs("id")%>">�޸�</a></td>
                  <td><a href="Bs_scType.asp?nbtyped=<%=rs("nbtyped")%>&amp;no=yes">ɾ��</a></td>
                </tr>
<%
rs.MoveNext 
loop 
%>
              </table>
<%
Set rs = Nothing 
Set Conn = Nothing 
%>
          </form>
		</td>
	</tr>
</table>
<%htmlend%>
