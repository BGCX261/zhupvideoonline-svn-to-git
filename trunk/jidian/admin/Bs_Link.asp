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
Response.Write("<script language=""JavaScript"">alert(""错误：您没输入网站名称，请返回检查！！"");history.go(-1);</script>")
response.end
end if

If note=""  or note="http://"Then
Response.Write("<script language=""JavaScript"">alert(""错误：您没输入图标连接，请返回检查！！"");history.go(-1);</script>")
response.end
end if


If link="" or link="http://" Then
Response.Write("<script language=""JavaScript"">alert(""错误：您没有输入超连接，请返回检查！！"");history.go(-1);</script>")
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
		<td class="a1">友情链接管理</td>
	</tr>
	<tr class="a4">
		<td>
          <form method="post" action="Bs_Link.asp?no=eshop">
			<table width="100%">
                <tr> 
                  <td class="iptlab"> 网站名称</td>
                  <td class="ipt"><input type="text" name="name" size="35" maxlength="40"></td>
                </tr>
                <tr> 
                  <td class="iptlab">网站说明</td>
                  <td class="ipt"><input type="text" name="note" size="50" maxlength="120"></td>
                </tr>
                <tr> 
                  <td class="iptlab"> 连接地址</td>
                  <td class="ipt"><input type="text" name="link" size="40" maxlength="50" value="http://"></td>
                </tr>
                <tr> 
                  <td colspan="2" class="iptsubmit">  
                      <input type="submit" name="Submit" value="提交">
                      <input type="reset" name="Submit2" value="重置">
                    </td>
                </tr>
              </table>
              <br> 
            <table width="100%">
                <tr class="opttitle"> 
                  <td>网站名称</td>
                  <td>网站说明</td>
                  <td>加入时间</td>
                  <td>操作</td>
                  <td>操作</td>
                </tr>
                <% do while not Rs.eof %>
                <tr class="opt"> 
                  <td><a href="<%=Rs("link")%>" target="_blank"><%=rs("name")%></a></td>
                  <td><%=Rs("note")%></td>
                  <td><%= FormatDateTime(rs("time"),2) %></td>
                  <td><a href="Bs_link_edit.asp?id=<%=Rs("id")%>">修改</a></td>
                  <td><a href="Bs_Link.asp?id=<%=Rs("id")%>&amp;no=yes">删除</a></td>
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

          <div>首页最多可显示10条友情链接</div>
		</td>
	</tr>
</table>
<%htmlend%>
