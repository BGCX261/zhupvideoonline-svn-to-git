<!--#include file="bsconfig.asp"-->

<%
dim Oldnbtyped
if Request.QueryString("no")="eshop" then

id=request("id")
nbtyped=request("nbtyped")
Oldnbtyped=request("Oldnbtyped")
Brief=request("Brief")


If nbtyped="" Then
response.write "SORRY <br>"
response.write "������!"
response.end
end if



Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from scType where id="&id
rs.open sql,conn,1,3

rs("nbtyped")=nbtyped
rs("Brief")=Brief
rs.update
rs.close
if nbtyped<>Oldnbtyped then
	conn.execute "Update Scientific set nbtyped='" & nbtyped & "' where nbtyped='" & Oldnbtyped & "'"
end if	
response.redirect "Bs_scType.asp"
end if
%>
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From scType where id="&id, conn,3,3
%>

<!-- #include file="Inc/Head.asp" -->
<table align="center" class="a2">
	<tr>
		<td class="a1">�޸Ľ�ѧ�������</td>
	</tr>
	<tr class="a4">
		<td>
          <form method="post" action="Bs_scType_edit.asp?no=eshop">
            <input type=hidden name=id value=<%=id%>>
			<input name="Oldnbtyped" type="hidden" id="Oldnbtyped" value="<%=rs("nbtyped")%>"> 
			<table width="100%">
                <tr> 
                  <td class="iptlab"> �޸��������</td>
                  <td class="ipt"><input type="text" name="nbtyped" class="inputtext" id="nbtyped" value="<%=rs("nbtyped")%>" size="35" maxlength="40"></td>
               </tr>
               <tr> 
                  <td class="iptlab"> ����˵��</td>
                  <td class="ipt"><textarea name="Brief" cols="50" rows="4"><%=rs("Brief")%></textarea></td>
                </tr>
                <tr> 
                  <td colspan="2" class="iptsubmit">  
					<input name="submit" type="submit" value="ȷ��">
					<input name="reset" type="reset" value="ȡ��">
				  </td>
                </tr>
              </table>
          </form>
		</td>
	</tr>
</table>
<%htmlend%>