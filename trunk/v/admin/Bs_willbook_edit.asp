<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->
<%if Request.QueryString("no")="eshop" then

id=request("id")
title=request("title")
publish=request("publish")
nbtyped=request("nbtyped")
content=Request("content")
IncludePic=Request("IncludePic")
DefaultPicUrl=Request("DefaultPicUrl")
UploadFiles=Request("UploadFiles")


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
sql="select * from newbook where id="&id
rs.open sql,conn,1,3

rs("title")=title
rs("publish")=publish
rs("nbtyped")=nbtyped
rs("content")=ubbcode(content)
if IncludePic="yes" then
		rs("IncludePic")=True
	else
		rs("IncludePic")=False
	end if
rs("DefaultPicUrl")=DefaultPicUrl
rs("UploadFiles")=UploadFiles
rs.update
rs.close
response.redirect "Bs_willbook.asp"
end if
%>
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From newbook where id="&id, conn,3,3
%>
<script>
function CheckForm()
{
  if (editor.EditMode.checked==true)
	  document.myform.content.value=editor.HtmlEdit.document.body.innerText;
  else
	  document.myform.content.value=editor.HtmlEdit.document.body.innerHTML; 
  if (document.myform.content.value=="")
  {
    alert("内容不能为空！");
	editor.HtmlEdit.focus();
	return false;
  }
  if (document.myform.title.value=="")
  {
    alert("标题不能为空！");
	myform.title.focus();
	return false;
  }
  if (document.myform.typed.value=="")
  {
    alert("类型不能为空！");
	myform.typed.focus();
	return false;
  }
  return true;  
}
function loadForm()
{
  editor.HtmlEdit.document.body.innerHTML=document.myform.content.value;
  return true
}
</script>
<!-- #include file="Inc/Head.asp" -->
<body onLoad="javascipt:setTimeout('loadForm()',1000);">
<table align="center" class="a2">
	<tr> 
		<td class="a1">修 改 视 频 教 程</td>
	</tr>
	<tr class="a4">
		<td>
					<form method="post" action="Bs_willbook_edit.asp?no=eshop" name="myform" onSubmit="return CheckForm();">
						<input type=hidden name=id value=<%=id%>>
							<table width="100%">
								<tr> 
									<td class="iptlab">标题</td>
									<td class="ipt"><input type="text" name="title" maxlength="30" size="25" value="<%=rs("title")%>"></td>
								</tr>
								<tr> 
									<td class="iptlab">来源</td>
									<td class="ipt"><input type="text" name="publish" maxlength="50" size="35" value="<%=rs("publish")%>"></td>
								</tr>
								<tr> 
									<td class="iptlab">类别</td>
									<td class="ipt">
<select  name="nbtyped">
              <%               
dim rs1,sql1,sel      
 set rs1=server.createobject("adodb.recordset")      
  sql1="select * from newbooktype"      
 rs1.open sql1,conn,1,1      
			  do while not rs1.eof      
                                sel="selected"          
		             response.write "<option  value='"+CStr(rs1("nbtyped"))+"' "         
		             if rs("nbtyped")=rs1("nbtyped") then Response.Write sel      
		             Response.Write ">"+rs1("nbtyped")+"</option>"+chr(13)+chr(10)      
		             rs1.movenext      
    		          loop      
			rs1.close      
			        %>
              </select>
</td>
								</tr>
								<tr> 
									<td class="iptlab">内容</td>
									<td class="ipt">
                                    <textarea name="content" cols="80" rows="15" id="content" style="display:none">
<%if rs("html")=false then
content=replace(rs("content"),"<br>",chr(13))
content=replace(content,"&nbsp;"," ")
else
content=rs("content")
end if
response.write content%></textarea>
<iframe ID="editor" src="../editor.asp" frameborder=1 scrolling=no width="700" height="405"></iframe>
</td>
								</tr>
								<tr>
								  <td>图片：</td>
								  <td><input name="IncludePic" type="hidden" id="IncludePic" value="yes">
								<input name="DefaultPicUrl" type="text" id="DefaultPicUrl" value="<%=rs("DefaultPicUrl")%>" size="50" maxlength="50">
								 
<select name="DefaultPicList" id="DefaultPicList" onChange="DefaultPicUrl.value=this.value;">
<option value="img/nopic.jpg" <% if rs("DefaultPicUrl")="" then response.write "selected" %>>不指定首页图片</option>
<%
if rs("UploadFiles")<>"" then
	dim IsOtherUrl
	IsOtherUrl=True
	if instr(rs("UploadFiles"),"|")>1 then
		dim arrUploadFiles,intTemp
		arrUploadFiles=split(rs("UploadFiles"),"|")						
		for intTemp=0 to ubound(arrUploadFiles)
			if rs("DefaultPicUrl")=arrUploadFiles(intTemp) then
				response.write "<option value='" & arrUploadFiles(intTemp) & "' selected>" & arrUploadFiles(intTemp) & "</option>"
				IsOtherUrl=False
			else
				response.write "<option value='" & arrUploadFiles(intTemp) & "'>" & arrUploadFiles(intTemp) & "</option>"
			end if
		next
	else
		if rs("UploadFiles")=rs("DefaultPicUrl") then
			response.write "<option value='" & rs("UploadFiles") & "' selected>" & rs("UploadFiles") & "</option>"
			IsOtherUrl=False
		else
			response.write "<option value='" & rs("UploadFiles") & "'>" & rs("UploadFiles") & "</option>"		
		end if
	end If
	if IsOtherUrl=True then
		response.write "<option value='" & rs("DefaultPicUrl") & "' selected>" & rs("DefaultPicUrl") & "</option>"
	end if
end if
%>
</select>
								<input name="UploadFiles" type="hidden" id="UploadFiles" value="<%=rs("UploadFiles")%>"></td>
								</tr>
								<tr> 
									<td colspan="2" class="iptsubmit"> 
											<input type="submit" value="确定">
											<input type="reset" value="从来">
										</td>
								</tr>
							</table>
					</form>
		</td>
	</tr>
</table>
<%htmlend%>
