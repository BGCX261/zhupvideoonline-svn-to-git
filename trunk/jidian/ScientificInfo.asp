<!--#include file="Inc/syscode.asp"-->
<!--#include file="Inc/eshopcode.asp"-->
<!-- #include file="Inc/Head.asp" -->
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
          <!-- #include file="L_teacher.asp" -->
      </td>
      <td class="bodyline"></td>
      <td class="bodyright">
      <!-- #include file="Inc/headAd.asp" -->
        <div class="mbox">
          <div class="positiontitle"><img src="img/title_ico.gif" /> ��ѧ����</div>
<%
if not isEmpty(request.QueryString("id")) then
id=request.QueryString("id")
else
id=1
end if

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From Scientific where id="&id, conn,3,3
'��¼���ʴ���
rs("counter")=rs("counter")+1
rs.update
nCounter=rs("counter")
'��������
content=rs("content")
%>
              <div class="contentbox">
                <div class="chptitletd"><%=rs("title")%></div>
                <div class="chphittd">�������:<%=rs("nbtyped")%> ����:[ <a href="javascript:ContentSize(16)">�� </a
><a href="javascript:ContentSize(14)">�� </a
><a href="javascript:ContentSize(12)">С </a
> ] ��������:[<%=rs("time")%>]  �Ķ�:[<%=rs("counter")%>]
                </div>
                <div class="contentview" id="BodyLabel">
				<%
				content=replace(rs("content"),"../","./")
				response.write content
				%>
				<%rs.close%>
                </div>
                <div class="prt">��<a href='javascript:window.print()'>��ӡ</a
>����<a href='javascript:history.back()'>����</a
>����<a href="javascript:window.scroll(0,-360)">����</a
>����<a href="javascript:self.close()">�ر�</a>��
                </div>
              </div>
        </div>
      </td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
