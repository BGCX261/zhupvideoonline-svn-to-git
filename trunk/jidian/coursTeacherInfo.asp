<!--#include file="Inc/syscode.asp"-->
<!--#include file="Inc/eshopcode.asp"-->
<!-- #include file="Inc/Head.asp" -->
<style>
.contentbox p{ text-indent:2em;}
</style>
<div id="body">
  <table class="bodytable">
    <tr>
      <td class="bodyleft">
          <!-- #include file="L_teacher.asp" -->
      </td>
      <td class="bodyline"></td>
      <td class="bodyright">
      <!-- #include file="Inc/headAd.asp" -->
        <div class="mbox">
            <div class="positiontitle"><img src="img/title_ico.gif" /> 教学管理 >> 教学团队</div>
<%
if not isEmpty(request.QueryString("id")) then
id=request.QueryString("id")
else
id=1
end if

set rs=server.createobject("adodb.recordset")
sqltext="select * from Teacher where sortId="&id
rs.open sqltext,conn,1,1
%>
              <div class="contentbox">
                  <div class="chptitletd"><%=rs("teachName")%></div>
                  <div class="contentview"><p><img src="<%=rs("Honor")%>" width="150" hspace="8" border="0" align="left" /><%=rs("teachInfo")%></p><p><%=rs("teachSay")%></p><p><%=rs("teachSubj")%><%rs.close%></p></div>
                  <div class="prt">【<a href='javascript:window.print()'>打印此页</a
>】【<a href='javascript:history.back()'>返回</a
>】【<a href="javascript:window.scroll(0,-360)">顶部</a
>】【<a href="javascript:self.close()">关闭</a>】</div>
              </div>
        </div>
      </td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
