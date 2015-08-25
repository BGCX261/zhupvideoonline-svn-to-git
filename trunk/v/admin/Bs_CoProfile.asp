<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->

<%Profile=Trim(Request("Profile")) %>
<%if Request.QueryString("no")="eshop" then
set rs=server.createobject("ADODB.Recordset")
sql="select * from main"
rs.open sql,conn,3,3
rs("Profile")=ubbcode(Profile)
rs.update
rs.close
response.redirect "Bs_CoProfile.asp"
end if
sql="select * from main"
Set rs_home= Server.CreateObject("ADODB.Recordset")
rs_home.open sql,conn,1,1
%>
<script>
function CheckForm()
{
  if (editor.EditMode.checked==true)
	  document.myform.Profile.value=editor.HtmlEdit.document.body.innerText;
  else
	  document.myform.Profile.value=editor.HtmlEdit.document.body.innerHTML; 
  if (document.myform.Profile.value=="")
  {
    alert("内容不能为空！");
	editor.HtmlEdit.focus();
	return false;
  }
  return true;  
}
setTimeout('loadForm()',1000); 
function loadForm()
{
  editor.HtmlEdit.document.body.innerHTML=document.myform.Profile.value;
  return true
}
</script>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("15")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">基 本 介 绍 设 置</td>
	</tr>
	<tr class="a4">
		<td>
			<form method="POST" action="Bs_CoProfile.asp?no=eshop" name="myform" onSubmit="return CheckForm();">
			<table width="100%">
				<tr> 
				<td class="iptlab">信息设置</td>
				<td class="ipt">
<textarea name="Profile" cols="80" rows="15" id="Profile" style="display:none">
<%
if rs_home("html")=false then
Profile=replace(rs_home("Profile"),"<br>",chr(13))
Profile=replace(Profile,"&nbsp;"," ")
else
Profile=rs_home("Profile")
end if
response.write Profile
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
<input type="submit" value=" 修 改 "	 name="cmdok"><input type="reset" value=" 重 写 " name="cmdcancel"></td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
</table>
<%htmlend%>
