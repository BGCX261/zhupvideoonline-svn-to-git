<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->

<%Demoflv=Trim(Request("Demoflv")) %>
<%if Request.QueryString("no")="eshop" then
set rs=server.createobject("ADODB.Recordset")
sql="select * from main"
rs.open sql,conn,3,3
rs("Demoflv")=ubbcode(Demoflv)
rs.update
rs.close
response.redirect "Bs_Demoflv.asp"
end if
sql="select * from main"
Set rs_home= Server.CreateObject("ADODB.Recordset")
rs_home.open sql,conn,1,1
%>
<script>
function CheckForm()
{
  if (editor.EditMode.checked==true)
	  document.myform.Demoflv.value=editor.HtmlEdit.document.body.innerText;
  else
	  document.myform.Demoflv.value=editor.HtmlEdit.document.body.innerHTML; 
  if (document.myform.Demoflv.value=="")
  {
    alert("�������ݲ���Ϊ�գ�");
	editor.HtmlEdit.focus();
	return false;
  }
  return true;  
}
setTimeout('loadForm()',1000); 
function loadForm()
{
  editor.HtmlEdit.document.body.innerHTML=document.myform.Demoflv.value;
  return true
}
</script>
<!-- #include file="Inc/Head.asp" -->
<table align="center" class="a2">
	<tr>
		<td class="a1">ʹ �� �� ʾ �� �� �� ��</td>
	</tr>
	<tr class="a4">
		<td>
			<form method="POST" action="Bs_Demoflv.asp?no=eshop"  name="myform" onSubmit="return CheckForm();">
			<table width="100%">
				<tr> 
				<td class="iptlab">��Ϣ����</td>
				<td class="ipt">
<textarea name="Demoflv" cols="80" rows="15" id="Demoflv" style="display:none">
<%
if rs_home("html")=false then
Demoflv=replace(rs_home("Demoflv"),"<br>",chr(13))
Demoflv=replace(Demoflv,"&nbsp;"," ")
else
Demoflv=rs_home("Demoflv")
end if
response.write Demoflv
%>
</textarea>
<iframe ID="editor" src="../editor1.asp" frameborder=1 scrolling=no width="700" height="405"></iframe>
</td>
				</tr>
				<tr> 
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
				<tr> 
					<td colspan="2" class="iptsubmit">
<input type="submit" value=" �� �� "	 name="cmdok"><input type="reset" value=" �� д " name="cmdcancel"></td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
</table>
<%htmlend%>
