<!--#include file="bsconfig.asp"-->

<%
dim Oldnbtyped
if Request.QueryString("no")="eshop" then

id=request("id")
nbtyped=request("nbtyped")
Oldnbtyped=request("Oldnbtyped")


If nbtyped="" Then
response.write "SORRY <br>"
response.write "请输入!"
response.end
end if



Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from newbooktype where id="&id
rs.open sql,conn,1,3

rs("nbtyped")=nbtyped
rs.update
rs.close
if nbtyped<>Oldnbtyped then
	conn.execute "Update newbook set nbtyped='" & nbtyped & "' where nbtyped='" & Oldnbtyped & "'"
end if	
response.redirect "Bs_nbType.asp"
end if
%>
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From newbooktype where id="&id, conn,3,3
%>

<!-- #include file="Inc/Head.asp" -->
<table align="center" class="a2">
	<tr> 
		<td class="a1">修 改 视 频 教 程 类 别</td>
	</tr>
	<tr class="a4">
		<td>
          <form method="post" action="Bs_nbType_edit.asp?no=eshop">
            <input type=hidden name=id value=<%=id%>>
			<input name="Oldnbtyped" type="hidden" id="Oldnbtyped" value="<%=rs("nbtyped")%>"> 
			<table width="100%">
                <tr> 
                  <td class="iptlab"> 修改类别名称</td>
                  <td class="ipt"><input type="text" name="nbtyped" class="inputtext" id="nbtyped" value="<%=rs("nbtyped")%>" size="35" maxlength="40"></td>
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
