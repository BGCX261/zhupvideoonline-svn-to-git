<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->

<%if Request.QueryString("no")="eshop" then

title=request("title")
siteUrl=request("siteUrl")
content=Request("content")


If title="" Then
response.write "SORRY <br>"
response.write "������������ݵı���!!<a href=""javascript:history.go(-1)"">��������</a>"
response.end
end if
If content="" Then
response.write "SORRY <br>"
response.write "���ݲ���Ϊ��!!<a href=""javascript:history.go(-1)"">��������</a>"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from ProSet "
rs.open sql,conn,1,3
rs.addnew
rs("title")=title
rs("siteUrl")=siteUrl
rs("content")=ubbcode(content)
rs("counter")=0
rs.update
rs.close
response.redirect "Bs_proSet.asp"
end if
%>
<script>
function CheckForm(){
  if (document.myform.content.value=="")
  {
    alert("���ݲ���Ϊ�գ�");
	return false;
  }
  return true;  
}
</script>
<!-- #include file="Inc/Head.asp" -->
<table align="center" class="a2">
	<tr>
		<td class="a1">���רҵ������Ϣ</td>
	</tr>
	<tr class="a4">
		<td>
          <form method="post" action="Bs_proSet_Add.asp?no=eshop" name="myform" onSubmit="return CheckForm()">
			<table width="100%">
								<tr> 
									<td class="iptlab">����</td>
									<td class="ipt"><input type="text" name="title" class="inputtext" maxlength="30" size="25"></td>
								</tr>
                                <tr> 
									<td class="iptlab">רҵ��ַ</td>
									<td class="ipt"><input name="siteUrl" type="text" class="inputtext" value="http://"  size="32"></td>
								</tr>
								<tr> 
									<td class="iptlab">����</td>
									<td class="ipt">
                                    <textarea name="content" cols="80" rows="15" id="content" style=" width:770px;">
									<%if rs("html")=false then
content=replace(rs("content"),"<br>",chr(13))
content=replace(content,"&nbsp;"," ")
else
content=rs("content")
end if
response.write content%></textarea>
</td>
								</tr>
								<tr> 
									<td colspan="2" class="iptsubmit">
											<input type="submit" value="ȷ��">
											<input type="reset" value="ȡ��">
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
<script type="text/javascript">$('#content').xheditor({upLinkUrl:"upload.asp?immediate=1",upLinkExt:"zip,rar,txt",upImgUrl:"upload.asp?immediate=1",upImgExt:"jpg,jpeg,gif,png",upFlashUrl:"upload.asp?immediate=1",upFlashExt:"swf",upMediaUrl:"upload.asp?immediate=1",upMediaExt:"avi"/*,�Ƿ�ʹ��ubb����beforeSetSource:ubb2html,beforeGetSource:html2ubb*/});</script>
<%htmlend%>
