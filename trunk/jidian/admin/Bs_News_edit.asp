<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->

<%if Request.QueryString("no")="eshop" then

id=request("id")
title=request("title")
typed=request("typed")
content=Trim(Request("content"))
author=request("author")
toTop=trim(request.form("toTop"))



If title="" Then
response.write "SORRY <br>"
response.write "请输入更新内容的标题"
response.end
end if
If content="" Then
response.write "SORRY <br>"
response.write "内容不能为空"
response.end
end if


Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from news where id="&id
rs.open sql,conn,1,3

rs("title")=title
rs("typed")=typed
rs("content")=ubbcode(content)
rs("mdy_time")=now()
rs("author")=author
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
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From news where id="&id, conn,3,3
%>
<script>
function CheckForm(){
  if (document.myform.content.value==""){
    alert("内容不能为空！");
	return false;
  }
  if (document.myform.typed.value==""){
    alert("类型不能为空！");
	myform.typed.focus();
	return false;
  }
  return true;  
}
</script>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("36")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">修改文章</td>
	</tr>
	<tr class="a4">
		<td>
			<form method="post" action="Bs_News_edit.asp?no=eshop" name="myform" onSubmit="return CheckForm()">
			<input type=hidden name=id value=<%=id%>>
			<table width="100%">
								<tr> 
									<td class="iptlab">标题</td>
									<td class="ipt"><input type="text" name="title" maxlength="100" size="25" value="<%=rs("title")%>">
									类别<select  name="typed">
              <%               
dim rs1,sql1,sel      
 set rs1=server.createobject("adodb.recordset")      
  sql1="select * from newsType"      
 rs1.open sql1,conn,1,1      
			  do while not rs1.eof      
                                sel="selected"          
		             response.write "<option  value='"+CStr(rs1("typed"))+"' "         
		             if rs("typed")=rs1("typed") then Response.Write sel      
		             Response.Write ">"+rs1("typed")+"</option>"+chr(13)+chr(10)      
		             rs1.movenext      
    		          loop      
			rs1.close      
			        %>
              </select>  作者<input name="author" type="text" id="author" value='<%if session("realname") then response.write(session("realname")) else response.write ("请输入")end if%>' />  是否置顶：<input name="toTop" type="checkbox" id="toTop" value="yes" <% if rs("toTop")=true then response.Write("checked") end if%>>
								是<span class="red">（如果选中将加粗红色显示）</span></td>
								</tr>
								<tr> 
									<td class="iptlab">内容</td>
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
											<input type="submit" value="确定">
											<input type="reset" value="从来">										</td>
								</tr>
							</table>
			</form>
		</td>
	</tr>
</table>
<script type="text/javascript" src="../xheditor/jquery.js"></script>
<script type="text/javascript" src="../xheditor/xheditor.js"></script>
<script type="text/javascript" src="../xheditor/xheditor_plugins/ubb.min.js"></script>
<script type="text/javascript">$('#content').xheditor({upLinkUrl:"upload.asp?immediate=1",upLinkExt:"zip,rar,txt",upImgUrl:"upload.asp?immediate=1",upImgExt:"jpg,jpeg,gif,png",upFlashUrl:"upload.asp?immediate=1",upFlashExt:"swf",upMediaUrl:"upload.asp?immediate=1",upMediaExt:"avi",internalScript:true/*,是否使用ubb编码beforeSetSource:ubb2html,beforeGetSource:html2ubb*/});</script>
<%htmlend%>
