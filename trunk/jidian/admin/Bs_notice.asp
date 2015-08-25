<!--#include file="bsconfig.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->

<%notice=Trim(Request("notice")) %>
<%noticeTime=Request("noticeTime")%>
<%sent=Request("sent")%>
<%if Request.QueryString("no")="eshop" then
set rs=server.createobject("adodb.recordset")
sql="select * from main"
rs.open sql,conn,3,3
rs("notice")=ubbcode(notice)
rs("noticeTime")=noticeTime
rs("sent")=sent
rs.update
rs.close
response.redirect "Bs_notice.asp"
end if
sql="select * from main"
Set rs_home= Server.CreateObject("ADODB.Recordset")
rs_home.open sql,conn,1,1
%>
<script>
function CheckForm(){
  if (document.myform.notice.value=="")
  {
    alert("内容不能为空！");
	return false;
  }
  return true;  
}
</script>
<!-- #include file="Inc/Head.asp" -->
<table align="center" class="a2">
	<tr>
		<td class="a1">操作帮助</td>
	</tr>
	<tr class="a4">
		<td>
          <form method="POST" action="Bs_notice.asp?no=eshop" name="myform" onSubmit="return CheckForm()">
			<table width="100%">
                            <tr> 
									<td class="iptlab"> </td>
									<td class="ipt"><input name="noticeTime" type="hidden" id="noticeTime" value="<%=now()%>" />
								    <input name="sent" type="hidden" id="sent" value="<%=session("RealName")%>" /></td>
                            </tr>
							<tr> 
								<td class="iptlab">信息设置</td>
								<td class="ipt"> <textarea name="notice" cols="80" rows="15" id="notice" style=" width:770px;">
<%
if rs_home("html")=false then
notice=replace(rs_home("notice"),"<br>",chr(13))
notice=replace(notice,"&nbsp;"," ")
else
notice=rs_home("notice")
end if
response.write notice
%>
</textarea>
</td>
							</tr> 
                             
                             <tr style="<%call checkdisplay("07")%>"> 
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
<script type="text/javascript">$('#notice').xheditor({upLinkUrl:"upload.asp?immediate=1",upLinkExt:"zip,rar,txt",upImgUrl:"upload.asp?immediate=1",upImgExt:"jpg,jpeg,gif,png",upFlashUrl:"upload.asp?immediate=1",upFlashExt:"swf",upMediaUrl:"upload.asp?immediate=1",upMediaExt:"avi"/*,是否使用ubb编码beforeSetSource:ubb2html,beforeGetSource:html2ubb*/});</script>
<%htmlend%>
