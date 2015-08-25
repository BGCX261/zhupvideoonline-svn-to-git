<!--#include file="bsconfig.asp"-->

<%
set rs=server.CreateObject("adodb.recordset")
rs.open "Select * From FacResourceType",conn,1,3

if Request.QueryString("no")="yes" then
frtype= Trim(Request("frtype"))
Brief= Trim(Request("Brief"))
if frtype<>"" then
sql="delete from FacResourceType where frtype='" & frtype & "'"
conn.Execute  sql
sql="delete from FacResource where frtype='" & frtype & "'"
conn.Execute sql
end if


Set rs= Nothing
Set Conn = Nothing
Response.Redirect "Bs_facType.asp"
end if



if Request.QueryString("no")="eshop" then
frtype=request.form("frtype")
Brief=request.form("Brief")

If frtype="" Then
response.write "SORRY <br>"
response.write "请输入!"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from FacResourceType "
rs.open sql,conn,1,3
rs.addnew
rs("frtype")=frtype
rs("Brief")=Brief

rs.update
Response.Redirect "Bs_facType.asp"
end if
%>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("58")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">类别管理</td>
	</tr>
	<tr class="a4">
		<td>
          <form method="post" action="Bs_facType.asp?no=eshop">
			<table width="100%">
                <tr> 
                  <td class="iptlab"> 新增类别名称</td>
                  <td class="ipt"><input type="text" name="frtype" size="35" maxlength="40"></td>
                </tr>
                <tr> 
                  <td class="iptlab"> 中文说明</td>
                  <td class="ipt"><textarea name="Brief" cols="50" rows="4"></textarea></td>
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
                  <td><%=rs("frtype")%></td>
                  <td><a href="Bs_facType_edit.asp?id=<%=rs("id")%>">修改</a></td>
                  <td><a href="Bs_facType.asp?frtype=<%=rs("frtype")%>&amp;no=yes">删除</a></td>
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
