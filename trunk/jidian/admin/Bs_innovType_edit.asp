<!--#include file="bsconfig.asp"-->

<%
dim Oldfrtype
if Request.QueryString("no")="eshop" then

id=request("id")
frtype=request("frtype")
Oldfrtype=request("Oldfrtype")
Brief=request("Brief")


If frtype="" Then
response.write "SORRY <br>"
response.write "请输入!"
response.end
end if



Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from innovType where id="&id
rs.open sql,conn,1,3

rs("frtype")=frtype
rs("Brief")=Brief
rs.update
rs.close
if frtype<>Oldfrtype then
	conn.execute "Update innovation set frtype='" & frtype & "' where frtype='" & Oldfrtype & "'"
end if	
response.redirect "Bs_innovType.asp"
end if
%>
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From innovType where id="&id, conn,3,3
%>

<!-- #include file="Inc/Head.asp" -->
<table align="center" class="a2">
	<tr>
		<td class="a1">修改类别</td>
	</tr>
	<tr class="a4">
		<td>
          <form method="post" action="Bs_innovType_edit.asp?no=eshop">
            <input type=hidden name=id value=<%=id%>>
			<input name="Oldfrtype" type="hidden" id="Oldfrtype" value="<%=rs("frtype")%>"> 
			<table width="100%">
                <tr> 
                  <td class="iptlab"> 修改类别名称</td>
                  <td class="ipt"><input type="text" name="frtype" class="inputtext" id="frtype" value="<%=rs("frtype")%>" size="35" maxlength="40"></td>
               </tr>
               <tr> 
                  <td class="iptlab"> 中文说明</td>
                  <td class="ipt"><textarea name="Brief" cols="50" rows="4"><%=rs("Brief")%></textarea></td>
                </tr>
                <tr> 
                  <td colspan="2" class="iptsubmit">  
					<input name="submit" type="submit" value="确定">
					<input name="reset" type="reset" value="取消">
				  </td>
                </tr>
              </table>
          </form>
		</td>
	</tr>
</table>
<%htmlend%>
