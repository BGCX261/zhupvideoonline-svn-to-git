<!--#include file="bsconfig.asp"-->

<%
dim Oldnbtyped
if Request.QueryString("no")="eshop" then

id=request("id")
nbtyped=request("nbtyped")
Oldnbtyped=request("Oldnbtyped")
Brief=request("Brief")


If nbtyped="" Then
response.write "SORRY <br>"
response.write "请输入!"
response.end
end if



Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from SubjCType where id="&id
rs.open sql,conn,1,3

rs("nbtyped")=nbtyped
rs("Brief")=Brief
rs.update
rs.close
if nbtyped<>Oldnbtyped then
	conn.execute "Update SubjConstruct set nbtyped='" & nbtyped & "' where nbtyped='" & Oldnbtyped & "'"
end if	
response.redirect "Bs_SubjCType.asp"
end if
%>
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From SubjCType where id="&id, conn,3,3
%>
<script>
function CheckForm(){
  if (document.myform.Brief.value=="")
  {
    alert("内容不能为空！");
	return false;
  }
  if (document.myform.typed.value=="")
  {
    alert("类型不能为空！");
	document.myform.typed.focus();
	return false;
  }
  return true;  
}
</script> 
<!-- #include file="Inc/Head.asp" -->
<table align="center" class="a2">
	<tr>
		<td class="a1">修改类别</td>
	</tr>
	<tr class="a4">
		<td>
          <form method="post" action="Bs_SubjCType_edit.asp?no=eshop" name="myform" onSubmit="return CheckForm()">
            <input type=hidden name=id value=<%=id%>>
			<input name="Oldnbtyped" type="hidden" id="Oldnbtyped" value="<%=rs("nbtyped")%>"> 
			<table width="100%">
                <tr> 
                  <td class="iptlab"> 修改类别</td>
                  <td class="ipt"><input type="text" name="nbtyped" class="inputtext" id="nbtyped" value="<%=rs("nbtyped")%>" size="35" maxlength="40"></td>
               </tr>
               <tr> 
                  <td class="iptlab"> 中文说明</td>
                  <td class="ipt"><textarea name="Brief" cols="80" rows="15" id="Brief" style=" width:770px;">
<%if rs("html")=false then
Brief=replace(rs("Brief"),"<br>",chr(13))
Brief=replace(Brief,"&nbsp;"," ")
else
Brief=rs("Brief")
end if
response.write Brief%></textarea>
</td>
                </tr>
                <tr> 
                  <td colspan="2" class="iptsubmit">  
					<input name="submit" type="submit" value="确定">
					<input name="reset" type="reset" value="取消">
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
<script type="text/javascript">$('#Brief').xheditor({upLinkUrl:"upload.asp?immediate=1",upLinkExt:"zip,rar,txt",upImgUrl:"upload.asp?immediate=1",upImgExt:"jpg,jpeg,gif,png",upFlashUrl:"upload.asp?immediate=1",upFlashExt:"swf",upMediaUrl:"upload.asp?immediate=1",upMediaExt:"avi"/*,是否使用ubb编码beforeSetSource:ubb2html,beforeGetSource:html2ubb*/});</script>
<%htmlend%>
