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
      <td class="bodyleft"><div id="leftbg">
          <!-- #include file="L_willbook.asp" -->
        </div></td>
      <td class="bodyline"></td>
      <td class="bodyright"><table class="mbox">
          <tr>
            <td><div class="rtitle"><img src="img/title_ico.gif" /> �� Ƶ �� �� �� ��</div></td>
          </tr>
          <tr>
            <td><%
if not isEmpty(request.QueryString("id")) then
id=request.QueryString("id")
else
id=1
end if

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From newbook where id="&id, conn,3,3
'��¼���ʴ���
rs("counter")=rs("counter")+1
rs.update
nCounter=rs("counter")
'��������
content=rs("content")
%>
              <table width="100%">
                <tr>
                  <td class="chptitletd"><%=rs("title")%></td>
                </tr>
                <tr>
                  <td class="chphittd">����:[ <a href="javascript:ContentSize(16)">�� </a
><a href="javascript:ContentSize(14)">�� </a
><a href="javascript:ContentSize(12)">С </a
> ] ��������:[<%=rs("time")%>]  �Ķ�:[<%=rs("counter")%>]</td>
                </tr>
				<tr>
                  <td class="chpparttd">��Դ: <%=rs("publish")%></td>
                </tr>
				<tr>
                  <td class="chpparttd">�������: <%=rs("nbtyped")%></td>
                </tr>
                <tr>
                  <td class="chpviewtd" id="BodyLabel">&nbsp;&nbsp;&nbsp;&nbsp;<%=content%>
                    <%rs.close%></td>
                </tr>
                <tr>
                  <td class="prt">��<a href='javascript:window.print()'>��ӡ</a
>����<a href='javascript:history.back()'>����</a
>����<a href="javascript:window.scroll(0,-360)">����</a
>����<a href="javascript:self.close()">�ر�</a>��</td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
</body>
</html>
