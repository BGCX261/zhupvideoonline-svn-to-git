<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->

<%if Request.QueryString("no")="eshop" then

title=request("title")
typed=request("typed")
content=Request("content")
author=request("author")
toTop=trim(request.form("toTop"))


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
sql="select * from news "
rs.open sql,conn,1,3
rs.addnew
rs("title")=title
rs("typed")=typed
rs("content")=ubbcode(content)
rs("author")=author
rs("counter")=0
if toTop="yes" then
	rs("toTop")=True
else
	rs("toTop")=False
end if
rs.update
rs.close
response.redirect "Bs_News.asp"
end if
%>
<script>
function CheckForm(){
  if (document.myform.content.value=="")
  {
    alert("���ݲ���Ϊ�գ�");
	return false;
  }
  if (document.myform.typed.value=="")
  {
    alert("���Ͳ���Ϊ�գ�");
	myform.typed.focus();
	return false;
  }
  return true;  
}
</script>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("35")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">����¹���</td>
	</tr>
	<tr class="a4">
		<td>
          <form method="post" action="Bs_News_Add.asp?no=eshop" name="myform" onSubmit="return CheckForm()">
			<table width="100%">
								<tr> 
									<td class="iptlab">����</td>
									<td class="ipt"><input type="text" name="title" class="inputtext" maxlength="100" size="25">
									���<select name="typed">
          <%                
dim rs1,sql1,sel1           
 set rs1=server.createobject("adodb.recordset")           
  sql1="select * from newsType"           
 rs1.open sql1,conn,1,1           
			  do while not rs1.eof           
                                sel="selected"               
		             response.write "<option " & sel & " value='"+CStr(rs1("typed"))+"'>"+rs1("typed")+"</option>"+chr(13)+chr(10)           
		             rs1.movenext           
    		          loop           
			rs1.close
			set rs1=nothing           
			      %>
          <option selected="selected" value="">��ѡ�����</option>
        </select>  ����<input name="author" type="text" id="author" value='<%if session("realname") then response.write(session("realname")) else response.write ("������")end if%>' /> �Ƿ��ö���<input name="toTop" type="checkbox" id="toTop" value="yes">
								��<span class="red">�����ѡ�н��Ӵֺ�ɫ��ʾ��</span></td>
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
<script type="text/javascript">$('#content').xheditor({upLinkUrl:"upload.asp?immediate=1",upLinkExt:"zip,rar,txt",upImgUrl:"upload.asp?immediate=1",upImgExt:"jpg,jpeg,gif,png",upFlashUrl:"upload.asp?immediate=1",upFlashExt:"swf",upMediaUrl:"upload.asp?immediate=1",upMediaExt:"avi",internalScript:true/*,�Ƿ�ʹ��ubb����beforeSetSource:ubb2html,beforeGetSource:html2ubb*/});</script>
<%htmlend%>
