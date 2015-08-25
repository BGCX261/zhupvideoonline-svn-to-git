<!--#include file="bsconfig.asp"-->

<%
set rs=server.CreateObject("adodb.recordset")
rs.open "Select * From newbooktype",conn,1,3

if Request.QueryString("no")="yes" then
nbtyped= Trim(Request("nbtyped"))
if nbtyped<>"" then
sql="delete from newbooktype where nbtyped='" & nbtyped & "'"
conn.Execute  sql
sql="delete from newbook where nbtyped='" & nbtyped & "'"
conn.Execute sql
end if


Set rs= Nothing
Set Conn = Nothing
Response.Redirect "Bs_nbType.asp"
end if



if Request.QueryString("no")="eshop" then
nbtyped=request.form("nbtyped")

If nbtyped="" Then
response.write "SORRY <br>"
response.write "请输入!"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from newbooktype "
rs.open sql,conn,1,3
rs.addnew
rs("nbtyped")=nbtyped

rs.update
Response.Redirect "Bs_nbType.asp"
end if
%>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("22")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">视 频 教 程 类 别 管 理</td>
	</tr>
	<tr class="a4">
		<td>
          <form method="post" action="Bs_nbType.asp?no=eshop">
			<table width="100%">
                <tr> 
                  <td class="iptlab"> 新增类别名称</td>
                  <td class="ipt"><input type="text" name="nbtyped" size="35" maxlength="40"></td>
                </tr>
                <tr> 
                  <td colspan="2" class="iptsubmit">  
                      <input type="submit" name="Submit" value="确定">
                      <input type="reset" name="Submit2" value="重置">
                    </td>
                </tr>
              </table>
              <br> 
            <table width="100%">
                <tr class="opttitle"> 
                  <td>类别名称</td>
                  <td>操作</td>
                  <td>操作</td>
                </tr>
                <% do while not rs.eof %>
                <tr class="opt"> 
                  <td><%=rs("nbtyped")%></td>
                  <td><a href="Bs_nbType_edit.asp?id=<%=rs("id")%>">修改</a></td>
                  <td><a href="Bs_nbType.asp?nbtyped=<%=rs("nbtyped")%>&amp;no=yes">删除</a></td>
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
