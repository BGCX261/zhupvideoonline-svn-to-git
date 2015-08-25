<!--#include file="Inc/syscode.asp"-->
<!--#include file="Inc/eshopcode.asp"-->
<!-- #include file="Inc/redHead.asp" -->
<script type="text/javascript">
var currentpos,timer; 
function initializeScroll() { timer=setInterval("scrollwindow()",80);} 
function scrollclear(){clearInterval(timer);}
function scrollwindow() {currentpos=document.body.scrollTop;window.scroll(0,++currentpos);if (currentpos != document.body.scrollTop) sc();} 
document.onmousedown=scrollclear
document.ondblclick=initializeScroll
</script>
<script type="text/javascript">
function ContentSize(size)
{
	var obj=document.all.BodyLabel;
	obj.style.fontSize=size+"px";
}
</script>
<div id="body">
  <table class="bodytable">
    <tr>
      <td class="bodyleft">
          <!-- #include file="L_subjConstruct.asp" -->
      </td>
      <td class="bodyline"></td>
      <td class="bodyright">
      <!-- #include file="Inc/headAd.asp" -->
        <div class="mbox">
           
<%
if not isEmpty(request.QueryString("id")) then
id=request.QueryString("id")
else
id=1
end if

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From SubjConstruct where id="&id, conn,3,3

response.write  "<div class='artpath'><img src='img/title_ico.gif' /> <a href='SubjConstructs.asp'>党建工作</a> >> " & rs("nbtyped") & "</div>"

'纪录访问次数
rs("counter")=rs("counter")+1
rs.update
nCounter=rs("counter")
'定义内容
content=ubbcode(rs("content"))
%>
              <div class="contentbox">
                  <div class="chptitletd"><%=rs("title")%></div>
                  <div class="chphittd">字体大小:[ <a href="javascript:ContentSize(18)">大 </a
><a href="javascript:ContentSize(14)">中 </a
><a href="javascript:ContentSize(12)">小 </a
> ] 时间:[<%=formatdatetime(rs("time"),2)%>]  点击:[<%=rs("counter")%>]</div>
                  <div class="contentview" id="BodyLabel">
				  <%
                    content=replace(rs("content"),"../","./")
                    response.write content
                  %>
                  <%rs.close%>
                  </div>
                  <div class="prt">【<a href='javascript:window.print()'>打印</a
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
