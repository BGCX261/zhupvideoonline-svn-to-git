<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->
<%if Request.QueryString("no")="eshop" then

id=request("id")
title=request("title")
frtype=request("frtype")
content=Request("content")


If title="" Then
response.write "SORRY <br>"
response.write "������������ݵı���"
response.end
end if
If content="" Then
response.write "SORRY <br>"
response.write "���ݲ���Ϊ��"
response.end
end if


Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from FacResource where id="&id
rs.open sql,conn,1,3

rs("title")=title
rs("frtype")=frtype
rs("content")=ubbcode(content)
rs("mdy_time")=now()
rs.update
rs.close
response.redirect "Bs_facResource.asp"
end if
%>
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From FacResource where id="&id, conn,3,3
%>
<script>
function CheckForm(){
  if (document.myform.content.value=="")
  {
    alert("���ݲ���Ϊ�գ�");
	return false;
  }
  if (document.myform.title.value=="")
  {
    alert("���ⲻ��Ϊ�գ�");
	myform.title.focus();
	return false;
  }
  if (document.myform.frtype.value=="")
  {
    alert("���Ͳ���Ϊ�գ�");
	myform.frtype.focus();
	return false;
  }
  return true;  
}
</script>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("30")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">�޸�ʵ����ѧ��Ϣ</td>
	</tr>
	<tr class="a4">
		<td>
					<form method="post" action="Bs_facResource_edit.asp?no=eshop" name="myform" onSubmit="return CheckForm();">
						<input type=hidden name=id value=<%=id%>>
							<table width="100%">
								<tr> 
									<td class="iptlab">����</td>
									<td class="ipt"><input type="text" name="title" maxlength="30" size="25" value="<%=rs("title")%>"></td>
								</tr>
								<tr> 
									<td class="iptlab">���</td>
									<td class="ipt">
<select  name="frtype">
              <%               
dim rs1,sql1,sel      
 set rs1=server.createobject("adodb.recordset")      
  sql1="select * from FacResourceType"      
 rs1.open sql1,conn,1,1      
			  do while not rs1.eof      
                                sel="selected"          
		             response.write "<option  value='"+CStr(rs1("frtype"))+"' "         
		             if rs("frtype")=rs1("frtype") then Response.Write sel      
		             Response.Write ">"+rs1("frtype")+"</option>"+chr(13)+chr(10)      
		             rs1.movenext      
    		          loop      
			rs1.close      
			        %>
              </select>
</td>
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
											<input type="reset" value="����">
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
