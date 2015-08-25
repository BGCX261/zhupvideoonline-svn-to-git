<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->

<%if Request.QueryString("no")="eshop" then

title=request("title")
content=Request("content")


If title="" Then
response.write "SORRY <br>"
response.write "请输入更新内容的标题!!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If content="" Then
response.write "SORRY <br>"
response.write "内容不能为空!!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from news "
rs.open sql,conn,1,3
rs.addnew
rs("title")=title
rs("content")=ubbcode(content)
rs("counter")=0
rs.update
rs.close
response.redirect "Bs_News.asp"
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
  return true;  
}
function loadForm()
{
  editor.HtmlEdit.document.body.innerHTML=document.myform.content.value;
  return true
}
</script>
<script>   
setTimeout('loadForm()',1000);
</script> 
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("30")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">添 加 新 闻 动 态</td>
	</tr>
	<tr class="a4">
		<td>
          <form method="post" action="Bs_News_Add.asp?no=eshop" name="myform" onSubmit="return CheckForm()">
			<table width="100%">
								<tr> 
									<td class="iptlab">标题</td>
									<td class="ipt"><input type="text" name="title" class="inputtext" maxlength="30" size="25"></td>
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
<iframe ID="editor" src="../editor1.asp" frameborder=1 scrolling=no width="700" height="405"></iframe>
<!--<iframe ID="Editor" name="Editor" src="ubb/edit.htm?id=content" frameBorder="0" marginHeight="0" marginWidth="0" scrolling="No" style="height:320;width:100%"></iframe>-->
</td>
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
