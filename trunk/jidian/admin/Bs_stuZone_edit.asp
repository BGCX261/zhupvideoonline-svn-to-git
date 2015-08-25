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
DefaultPicUrl=Request("honor")
UploadFiles=Request("UploadFiles")
Elite=trim(request.form("Elite"))
Passed=trim(request.form("Passed"))
UpdateTime=trim(request.form("UpdateTime"))
author=request("author")

sqlBig="select * from StuZoneType where nbtyped='" & nbtyped & "'"
Set rsBig= Server.CreateObject("ADODB.Recordset")
rsBig.open sqlBig,conn,1,1
rsBig.close


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
sql="select * from StuZone where id="&id
rs.open sql,conn,1,3

rs("title")=title
rs("publish")=publish
rs("nbtyped")=nbtyped
rs("Passed")=Passed
rs("content")=ubbcode(content)
if UpdateTime<>"" and IsDate(UpdateTime)=true then
		UpdateTime=CDate(UpdateTime)
	else
		UpdateTime=now()
	end if
if IncludePic="yes" then
		rs("IncludePic")=True
	else
		rs("IncludePic")=False
	end if
if Elite="yes" then
		rs("Elite")=True
	else
		rs("Elite")=False
	end if
rs("UpdateTime")=UpdateTime
rs("DefaultPicUrl")=DefaultPicUrl
rs("UploadFiles")=UploadFiles
rs("author")=author
rs.update
rs.close
response.redirect "Bs_stuZone.asp"
end if
%>
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From StuZone where id="&id, conn,3,3
%>
<script>
function CheckForm(){
  if (document.myform.content.value==""){
    alert("内容不能为空！");
	return false;
  }
  if (document.myform.title.value==""){
    alert("标题不能为空！");
	myform.title.focus();
	return false;
  }
  if (document.myform.nbtyped.value==""){
    alert("类型不能为空！");
	myform.nbtyped.focus();
	return false;
  }
  return true;  
}
</script>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("45")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">修改文章</td>
	</tr>
	<tr class="a4">
		<td>
					<form method="post" action="Bs_stuZone_edit.asp?no=eshop" name="myform" onSubmit="return CheckForm();">
						<input type=hidden name=id value=<%=id%>>
							<table width="100%">
								<tr> 
									<td class="iptlab">标题</td>
									<td class="ipt"><input type="text" name="title" maxlength="30" size="25" value="<%=rs("title")%>"></td>
								</tr>
								<tr> 
									<td class="iptlab">摘要</td>
									<td class="ipt"><input type="text" name="publish" maxlength="50" size="35" value="<%=rs("publish")%>"></td>
								</tr>
								<tr> 
									<td class="iptlab">类别</td>
									<td class="ipt">
<select  name="nbtyped">
              <%               
dim rs1,sql1,sel      
 set rs1=server.createobject("adodb.recordset")      
  sql1="select * from StuZoneType"      
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
作者<input name="author" type="text" id="author" value='<%if session("realname") then response.write(session("realname")) else response.write ("请输入")end if%>' readonly="readonly" />
</td>
								</tr>
								<tr> 
									<td class="iptlab">内容</td>
									<td class="ipt">
                                    <textarea name="content" cols="80" rows="15" id="content" style="width:770px;">
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
								  <td>首页图片：</td>
								  <td><input name="IncludePic" type="hidden" id="IncludePic" value="yes">
								<input name="honor" type="text" id="honor" value="<%=rs("DefaultPicUrl")%>" size="50" maxlength="50">
                                <iframe name="ad" frameborder=0 width=400px height=25 scrolling=auto src=../uploadb.asp></iframe>
								 
<select name="DefaultPicList" id="DefaultPicList" onChange="honor.value=this.value;">
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
								<td class="iptlab">是否推荐：</td>
							  <td class="ipt">
								<input name="Elite" type="checkbox" id="Elite" value="yes" <% if rs("Elite")=true then response.Write("checked") end if%>>
								是<span class="red">（如果选中的话将在首页显示）</span></td>
							</tr>
                            <tr> 
								<td class="iptlab">是否审核：</td>
								<td class="ipt">
								
                                <input name="Passed" type="checkbox" id="Passed" value="1" <% if instr(manage,str)<=0 then response.Write("") else if rsSysAdmin("Passed")=1 then response.Write("checked") end if%> <%call checkverify("47")%>>
								是<span class="red">（如果选中的话将直接发布）</span></td>
							</tr>
							<tr> 
								<td class="iptlab">录入时间：</td>
								<td class="ipt">
								<input name="UpdateTime" type="text" id="UpdateTime" value="<%=now()%>" maxlength="50">
								当前时间为：<%=now()%> 注意不要改变格式。</td>
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
<script type="text/javascript" src="../xheditor/jquery.js"></script>
<script type="text/javascript" src="../xheditor/xheditor.js"></script>
<script type="text/javascript" src="../xheditor/xheditor_plugins/ubb.min.js"></script>
<script type="text/javascript">$('#content').xheditor({upLinkUrl:"upload.asp?immediate=1",upLinkExt:"zip,rar,txt",upImgUrl:"upload.asp?immediate=1",upImgExt:"jpg,jpeg,gif,png",upFlashUrl:"upload.asp?immediate=1",upFlashExt:"swf",upMediaUrl:"upload.asp?immediate=1",upMediaExt:"avi"/*,是否使用ubb编码beforeSetSource:ubb2html,beforeGetSource:html2ubb*/});</script>
<%htmlend%>
