<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->
<%if Request.QueryString("no")="eshop" then

title=request("title")
publish=request("publish")
nbtyped=request("nbtyped")
content=Request("content")
IncludePic=Request("IncludePic")
DefaultPicUrl=Request("DefaultPicUrl")
UploadFiles=Request("UploadFiles")


If title="" Then
response.write "SORRY <br>"
response.write "请输入标题!!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If content="" Then
response.write "SORRY <br>"
response.write "内容不能为空!!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from newbook "
rs.open sql,conn,1,3
rs.addnew
rs("title")=title
rs("publish")=publish
rs("nbtyped")=nbtyped
rs("content")=ubbcode(content)
rs("counter")=0
if IncludePic="yes" then
		rs("IncludePic")=True
	else
		rs("IncludePic")=False
	end if
'***************************************
	'删除无用的上传文件
	if ObjInstalled=True and UploadFiles<>"" then
		dim fso,strRubbishFile
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		if instr(UploadFiles,"|")>1 then
			dim arrUploadFiles,intTemp
			arrUploadFiles=split(UploadFiles,"|")
			UploadFiles=""
			for intTemp=0 to ubound(arrUploadFiles)
				if instr(Content,arrUploadFiles(intTemp))<=0 and arrUploadFiles(intTemp)<>DefaultPicUrl then
					strRubbishFile=server.MapPath("../" & arrUploadFiles(intTemp))
					if fso.FileExists(strRubbishFile) then
						fso.DeleteFile(strRubbishFile)
						response.write "<br><li>" & arrUploadFiles(intTemp) & "在文章中没有用到，也没有被设为首页图片，所以已经被删除！</li>"
					end if
				else
					if intTemp=0 then
						UploadFiles=arrUploadFiles(intTemp)
					else
						UploadFiles=UploadFiles & "|" & arrUploadFiles(intTemp)
					end if
				end if
			next
		else
			if instr(Content,UploadFiles)<=0 and UploadFiles<>DefaultPicUrl then
				strRubbishFile=server.MapPath("../" & UploadFiles)
				if fso.FileExists(strRubbishFile) then
					fso.DeleteFile(strRubbishFile)
					response.write "<br><li>" & UploadFiles & "在文章中没有用到，也没有被设为首页图片，所以已经被删除！</li>"
				end if
				UploadFiles=""
			end if
		end if
		set fso=nothing
	end If
	'结束
	'***************************************	
rs("DefaultPicUrl")=DefaultPicUrl
rs("UploadFiles")=UploadFiles
rs.update
rs.close
response.redirect "Bs_willbook.asp"
end if
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
  if (document.myform.nbtyped.value=="")
  {
    alert("类型不能为空！");
	myform.nbtyped.focus();
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
<%call checkmanage("26")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">添 加 视 频 教 程</td>
	</tr>
	<tr class="a4">
		<td>
          <form method="post" action="Bs_willbook_Add.asp?no=eshop" name="myform" onSubmit="return CheckForm();">
		<table width="100%">
								<tr> 
									<td class="iptlab"> 标题</td>
									<td class="ipt"><input type="text" name="title" class="inputtext" maxlength="30" size="25"></td>
								</tr>
								<tr> 
									<td class="iptlab"> 来源</td>
									<td class="ipt"><input type="text" name="publish" class="inputtext" maxlength="50" size="35"></td>
								</tr>
								<tr> 
									<td class="iptlab"> 类别</td>
									<td class="ipt">
<select name="nbtyped">
          <%                
dim rs1,sql1,sel1           
 set rs1=server.createobject("adodb.recordset")           
  sql1="select * from newbooktype"           
 rs1.open sql1,conn,1,1           
			  do while not rs1.eof           
                                sel="selected"               
		             response.write "<option " & sel & " value='"+CStr(rs1("nbtyped"))+"'>"+rs1("nbtyped")+"</option>"+chr(13)+chr(10)           
		             rs1.movenext           
    		          loop           
			rs1.close
			set rs1=nothing           
			      %>
          <option selected="selected" value="">请选择类别</option>
        </select>
</td>
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
								  <td> 
								<input name="IncludePic" type="hidden" id="IncludePic" value="yes">
								<input name="DefaultPicUrl" type="text" id="DefaultPicUrl" value="img/nopic.jpg" size="40" maxlength="120"> 
								<br>首页的图片,直接从上传图片中选择： 
								<select name="DefaultPicList" id="select" onChange="DefaultPicUrl.value=this.value;">
								<option selected value="img/nopic.jpg">不指定首页图片</option>
								</select> <input name="UploadFiles" type="hidden" id="UploadFiles"></td>
								</tr>
								<tr> 
									<td colspan="2" class="iptsubmit"> 
											<input type="submit" value="确定">
											<input type="reset" value="取消">
										</td>
								</tr>
							</table>
		</form>
		</td>
	</tr>
</table>
<%htmlend%>
