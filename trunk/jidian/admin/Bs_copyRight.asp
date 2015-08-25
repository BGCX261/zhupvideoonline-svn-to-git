<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->

<%copyRight=Trim(Request("copyRight")) %>
<%if Request.QueryString("no")="eshop" then
set rs=server.createobject("adodb.recordset")
sql="select * from syswork where code='101'"
rs.open sql,conn,3,3
rs("copyRight")=ubbcode(copyRight)
rs.update
rs.close
response.redirect "Bs_copyRight.asp"
end if
sql="select * from syswork where code='101'"
Set rs_home= Server.CreateObject("ADODB.Recordset")
rs_home.open sql,conn,1,1
%>
<script>
function CheckForm(){
  if (document.myform.copyRight.value=="")
  {
    alert("内容不能为空！");
	return false;
  }
  return true;  
}
</script>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("56")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">页脚版权设置</td>
	</tr>
	<tr class="a4">
		<td>
          <form method="POST" action="Bs_copyRight.asp?no=eshop" name="myform" onSubmit="return CheckForm()">
			<table width="100%">
							<tr> 
								<td class="iptlab">信息设置</td>
								<td class="ipt"> <textarea name="copyRight" cols="80" rows="15" id="copyRight" style="width:770px;">
<%
if rs_home("html")=false then
copyRight=replace(rs_home("copyRight"),"<br>",chr(13))
copyRight=replace(copyRight,"&nbsp;"," ")
else
copyRight=rs_home("copyRight")
end if
response.write copyRight
%>
</textarea>
</td>
							</tr> 
							<tr> 
								<td colspan="2" class="iptsubmit">
									<input type="submit" value=" 修 改 " name="cmdok">
									<input type="reset" value=" 重 写 " name="cmdcancel">
								</td>
							</tr>
						</table>
		  </form>
		</td>
	</tr>
</table>
<script type="text/javascript" src="../xheditor/jquery.js"></script>
<script type="text/javascript" src="../xheditor/xheditor.js"></script>
<script type="text/javascript" src="../xheditor/xheditor_plugins/ubb.min.js"></script>
<script type="text/javascript">$('#copyRight').xheditor({upLinkUrl:"upload.asp?immediate=1",upLinkExt:"zip,rar,txt",upImgUrl:"upload.asp?immediate=1",upImgExt:"jpg,jpeg,gif,png",upFlashUrl:"upload.asp?immediate=1",upFlashExt:"swf",upMediaUrl:"upload.asp?immediate=1",upMediaExt:"avi"/*,是否使用ubb编码beforeSetSource:ubb2html,beforeGetSource:html2ubb*/});</script>
<%htmlend%>
