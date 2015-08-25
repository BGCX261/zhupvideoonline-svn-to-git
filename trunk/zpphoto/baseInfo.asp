<!--#include file="security.asp"-->
<!--#include file="dbConn.asp"-->
<!--#include file="head.asp"-->
<body>
<div class="admCont"> 
  <div class="admTop"></div>
  <div class="admLeft">
      <!--#include file="left.asp"-->
  </div>
  <div class="admRight">
<%
   dim rs,sql,sel           
   set rs=server.createobject("adodb.recordset")
   sql="select * from siteUsers where id=1"
    rs.open sql,conn,1,1
    %>
 <table class="iptTable">
        <form method="POST" onsubmit action="saveBaseInfo.asp">
          <tr> 
            <td colspan="2" class="doTitle"> 超级用户</td>
          </tr>
          <tr>
            <td>用户：</td>
            <td nowrap><input type="text" name="name" size="16" value="<%=rs("name")%>"></td>
          </tr>
          <tr>
            <td>密码：</td>
            <td nowrap><input type="text" name="pwd" size="16" value="<%=rs("pwd")%>"></td>
          </tr>
          <tr> 
            <td width="40%">&nbsp;</td>
            <td width="40%" nowrap><font size="2">
              <input type="submit" value="修改" name="my">
              </font></td>
          </tr>
        </form>
        <tr> 
          <td class="doTitle" colspan="2">栏目管理</td>
          <%set rs=server.createobject("adodb.recordset")
   sql="select * from type"
   rs.open sql,conn,1,3
  do while not rs.eof%>
        </tr>
        <form method="POST" onsubmit action="saveBaseInfo.asp?typeid=<%=rs("typeid")%>">
          <tr> 
            <td width="20%" nowrap><%=rs("type")%></td>
            <td width="80%" nowrap> <input type="text" name="typename" size="16"> 
              <input type="submit" value="修改" name="change"> 
              <input type="submit" value="删除" name="delet"> 
            </td>
          </tr>
        </form>
        <% rs.movenext
      loop
      %>
        <form method="POST" onsubmit action="saveBaseInfo.asp">
          <tr bgcolor="#FFFFFF"> 
            <td nowrap>新增：</td>
            <td nowrap> <input type="text" name="type" size="16"> 
            <input type="submit" value="新增" name="addname"> </td>
          </tr>
        </form>
        <tr> 
          <td class="doTitle" colspan="2">网页信息</td>
        </tr>
        <%set rs=server.createobject("adodb.recordset")
  rs.open "select * from baseSet",conn,1,3
  %>
        <form method="POST" action="saveBaseInfo.asp">
          <tr> 
            <td width="20%"> 站点名称：</td>
            <td width="80%"> <input name="homes" maxlength="50" value="<%=rs("homes")%>"></td>
          </tr>
          <tr> 
            <td> 国际域名：</td>
            <td> <input name="url" maxlength="50" value="<%=rs("url")%>"> 
            </td>
          </tr>
          <tr>
            <td>网站声明：</td>
            <td><input name="announces" maxlength="50" value="<%=rs("announces")%>"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td><input type="submit" value=" 修 改 " name="cmdok">
              <input type="reset" value=" 重 写 " name="cmdcancel"></td>
          </tr>
        </form>
      </table> 
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
    </div>
</div>
</body>
</html>


