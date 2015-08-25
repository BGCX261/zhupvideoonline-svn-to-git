<!--#include file="bsconfig.asp"-->

<%if Request.QueryString("no")="eshop" then

id=request("id")
name=request("name")
note=request("note")
Enname=request("Enname")
Ennote=request("Ennote")
link=request("link")


If name="" Then
response.write "SORRY <br>"
response.write "请输入更新内容"
response.end
end if
'If Enname="" Then
'response.write "SORRY <br>"
'response.write "请输入更新内容"
'response.end
'end if
If note="" Then
response.write "SORRY <br>"
response.write "内容不能为空"
response.end
end if
'If Ennote="" Then
'response.write "SORRY <br>"
'response.write "内容不能为空"
'response.end
'end if
If link="" Then
response.write "SORRY <br>"
response.write "内容不能为空"
response.end
end if


Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from links where id="&id
rs.open sql,conn,1,3

rs("name")=name
rs("note")=note
'rs("Enname")=Enname
'rs("Ennote")=Ennote
rs("link")=link
rs.update
rs.close
response.redirect "Bs_Link.asp"
end if
%>
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From links where id="&id, conn,3,3
%>

<!-- #include file="Inc/Head.asp" -->
<table align="center" class="a2">
	<tr>
		<td class="a1">修改友情链接</td>
	</tr>
	<tr class="a4">
		<td>
          <form method="post" action="Bs_Link_edit.asp?no=eshop">
            <input type=hidden name=id value=<%=id%>>
			<table width="100%">
                <tr> 
                  <td class="iptlab"> 网站名称</td>
                  <td class="ipt"><input type="text" name="name" class="inputtext" id="name" value="<%=rs("name")%>" size="35" maxlength="40"></td>
               </tr>
			   <!--<tr> 
                  <td class="iptlab"> English名称</td>
                  <td class="ipt"><input type="text" name="Enname" class="inputtext" id="Enname" value="<%=rs("Enname")%>" size="35" maxlength="40"></td>
               </tr>-->
                <tr> 
                  <td class="iptlab">网站说明</td>
                  <td class="ipt"><input type="text" name="note" class="inputtext" id="note" value="<%=rs("note")%>" size="50" maxlength="120"></td>
                </tr>
				<!--<tr> 
                  <td class="iptlab">English说明</td>
                  <td class="ipt"><input type="text" name="Ennote" class="inputtext" id="Ennote" value="<%=rs("Ennote")%>" size="50" maxlength="120"></td>
                </tr>-->
                <tr> 
                  <td class="iptlab"> 连接地址</td>
                  <td class="ipt"><input type="text" name="link" class="inputtext" id="link" value="<%=rs("link")%>" size="40" maxlength="50"></td>
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
