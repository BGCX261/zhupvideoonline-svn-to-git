<!--#include file="bsconfig.asp"-->

<%
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = Conn
Rs.Open "SELECT * FROM links Order BY id"

if Request.QueryString("no")="yes" then
id= Trim(Request("id"))
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = Conn
Rs.Open "DELETE FROM links Where id="&id,Conn,2,3,1


Set Rs= Nothing
Set Conn = Nothing
Response.Redirect "Bs_Link.asp"
end if



if Request.QueryString("no")="eshop" then
note=request.form("note")
link=request.form("link")
name=request.form("name")

If name=""  or note="http://"Then
Response.Write("<script language=""JavaScript"">alert(""������û������վ���ƣ��뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if

If note=""  or note="http://"Then
Response.Write("<script language=""JavaScript"">alert(""������û����ͼ�����ӣ��뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if


If link="" or link="http://" Then
Response.Write("<script language=""JavaScript"">alert(""������û�����볬���ӣ��뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from links "
rs.open sql,conn,1,3
rs.addnew
rs("name")=name
rs("note")=note
rs("link")=link

rs.update
Response.Redirect "Bs_Link.asp"
end if
%>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("12")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">�������ӹ���</td>
	</tr>
	<tr class="a4">
		<td>
          <form method="post" action="Bs_Link.asp?no=eshop">
			<table width="100%">
                <tr> 
                  <td class="iptlab"> ��վ����</td>
                  <td class="ipt"><input type="text" name="name" size="35" maxlength="40"></td>
                </tr>
                <tr> 
                  <td class="iptlab">��վ˵��</td>
                  <td class="ipt"><input type="text" name="note" size="50" maxlength="120"></td>
                </tr>
                <tr> 
                  <td class="iptlab"> ���ӵ�ַ</td>
                  <td class="ipt"><input type="text" name="link" size="40" maxlength="50" value="http://"></td>
                </tr>
                <tr> 
                  <td colspan="2" class="iptsubmit">  
                      <input type="submit" name="Submit" value="�ύ">
                      <input type="reset" name="Submit2" value="����">
                    </td>
                </tr>
              </table>
              <br> 
            <table width="100%">
                <tr class="opttitle"> 
                  <td>��վ����</td>
                  <td>��վ˵��</td>
                  <td>����ʱ��</td>
                  <td>����</td>
                  <td>����</td>
                </tr>
                <% do while not Rs.eof %>
                <tr class="opt"> 
                  <td><a href="<%=Rs("link")%>" target="_blank"><%=rs("name")%></a></td>
                  <td><%=Rs("note")%></td>
                  <td><%= FormatDateTime(rs("time"),2) %></td>
                  <td><a href="Bs_link_edit.asp?id=<%=Rs("id")%>">�޸�</a></td>
                  <td><a href="Bs_Link.asp?id=<%=Rs("id")%>&amp;no=yes">ɾ��</a></td>
                </tr>
<%
Rs.MoveNext 
loop 
%>
              </table>
<%
Set Rs = Nothing 
Set Conn = Nothing 
%>
          </form>

          <div>��ҳ������ʾ10����������</div>
		</td>
	</tr>
</table>
<%htmlend%>
