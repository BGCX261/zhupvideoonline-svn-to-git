<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->

<%Declare=Trim(Request("Declare")) %>
<%if Request.QueryString("no")="eshop" then
set rs=server.createobject("adodb.recordset")
sql="select * from syswork"
rs.open sql,conn,3,3
rs("Declare")=ubbcode(Declare)
rs.update
rs.close
response.redirect "Bs_Declare.asp"
end if
sql="select * from syswork"
Set rs_home= Server.CreateObject("ADODB.Recordset")
rs_home.open sql,conn,1,1
%>
<script>
function CheckForm()
{
  if (editor.EditMode.checked==true)
	  document.myform.Declare.value=editor.HtmlEdit.document.body.innerText;
  else
	  document.myform.Declare.value=editor.HtmlEdit.document.body.innerHTML; 
  if (document.myform.Declare.value=="")
  {
    alert("内容不能为空！");
	editor.HtmlEdit.focus();
	return false;
  }
  return true;  
}
function loadForm()
{
  editor.HtmlEdit.document.body.innerHTML=document.myform.Declare.value;
  return true
}
</script>
<script>   
setTimeout('loadForm()',1000);
</script> 
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("14")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">网 站 声 明 设 置</td>
	</tr>
	<tr class="a4">
		<td>
			<form method="POST" action="Bs_Declare.asp?no=eshop" name="myform" onSubmit="return CheckForm();">
			<table width="100%">
				<tr> 
				<td class="iptlab"> 信息设置</td>
				<td class="ipt">
<textarea name="Declare" cols="80" rows="15" id="Declare" style="display:none">
<%
if rs_home("html")=false then
Declare=replace(rs_home("Declare"),"<br>",chr(13))
Declare=replace(Declare,"&nbsp;"," ")
else
Declare=rs_home("Declare")
end if
response.write Declare
%>
</textarea>
<iframe ID="editor" src="../editor1.asp" frameborder=1 scrolling=no width="700" height="405"></iframe></td>
				</tr>
				<tr> 
					<td colspan="2" class="iptsubmit">
<input type="submit" value=" 修 改 "	 name="cmdok"><input type="reset" value=" 重 写 " name="cmdcancel"></td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
</table>
<%htmlend%>
