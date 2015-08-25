<!--#include file="bsconfig.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->

<%if Request.QueryString("no")="eshop" then

id=request("id")
realname=request("realname")
adminRole=request("adminRole")
ubelong=Trim(Request("ubelong"))
uemail=request("uemail")
uphone=request("uphone")
uqq=Request("uqq")
admission=request("admission")
uprofile=Trim(Request("uprofile"))
ufaculty=Trim(Request("ufaculty"))


If realname="" Then
response.write "SORRY <br>"
response.write "请输入真实姓名"
response.end
end if
If uprofile="" Then
response.write "SORRY <br>"
response.write "内容不能为空"
response.end
end if


Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Bs_User where id="&id
rs.open sql,conn,1,3

rs("realname")=realname
rs("adminRole")=adminRole
rs("ubelong")=ubelong
rs("uemail")=uemail
rs("uphone")=uphone
rs("uqq")=uqq
rs("admission")=admission
rs("uprofile")=ubbcode(uprofile)
rs("ufaculty")=ufaculty
rs.update
rs.close
response.redirect "Bs_Admin_List.asp"
end if
%>
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From Bs_User where id="&id, conn,3,3
%>
<script>
function CheckForm()
{
  if (document.myform.uprofile.value=="")
  {
    alert("内容不能为空！");
	editor.HtmlEdit.focus();
	return false;
  }
  return true;  
}
</script>
<!-- #include file="Inc/Head.asp" -->
<table align="center" class="a2">
	<tr>
		<td class="a1">修 改 管 理 员 信 息</td>
	</tr>
	<tr class="a4">
		<td>
			<form method="post" action="Bs_Admin_edit.asp?no=eshop" name="myform" onSubmit="return CheckForm()">
			<input type=hidden name=id value=<%=id%>>
			<table width="100%">
								<tr> 
									<td class="iptlab">用户名</td>
									<td class="ipt"><input type="text" name="adminName" value="<%=rs("adminName")%>" readonly="readonly"></td>
                                    <td class="iptlab">真实姓名</td>
									<td class="ipt"><input name="realname" type="text" value="<%=rs("realname")%>" readonly="readonly"></td>
								</tr>
                                <tr> 
									<td class="iptlab">角色身份</td>
									<td class="ipt"><input type="text" name="adminRole" value="<%=rs("adminRole")%>" readonly="readonly" /></td>
                                    <td class="iptlab">专业班级</td>
									<td class="ipt"><input type="text" name="ubelong" value="<%=rs("ubelong")%>"></td>
								</tr>
                                <tr> 
									<td class="iptlab">电子邮件</td>
									<td class="ipt"><input type="text" name="uemail" value="<%=rs("uemail")%>"></td>
                                    <td class="iptlab">联系电话</td>
									<td class="ipt"><input type="text" name="uphone" value="<%=rs("uphone")%>"></td>
								</tr>
                                <tr> 
									<td class="iptlab">QQ</td>
									<td class="ipt"><input type="text" name="uqq" value="<%=rs("uqq")%>"></td>
                                    <td class="iptlab">入学年份</td>
									<td class="ipt"><input name="admission" type="text" value="<%=rs("admission")%>" readonly="readonly"></td>
								</tr>
								<tr> 
									<td class="iptlab">内容</td>
									<td class="ipt">
                                    <input name="uprofile" type="text" size="36" id="uprofile" value="<%=rs("uprofile")%>"></td>
                                    <td class="iptlab"> 所在院系部：</td>
                                    <td class="ipt">
                                    <select name="ufaculty">
                                      <option selected="selected" value="<%=rs("ufaculty")%>"><%=rs("ufaculty")%></option>
                                      <option value="工程与信息学院">工程与信息学院</option>
                                      <option value="经济管理学院">经济管理学院</option>
                                      <option value="国际合作与交流学院">国际合作与交流学院</option>
                                      <option value="人文社科系">人文社科系</option>
                                      <option value="艺术设计系">艺术设计系</option>
                                      <option value="旅游系">旅游系</option>
                                      <option value="工程与信息学院">工程与信息学院</option>
                                      <option value="思政部">思政部</option>
                                      <option value="成人教育学院">成人教育学院</option>
                                      <option value="行政或其它">行政或其它</option>
                                    </select>
                                    </td>
                                </tr>
								<tr> 
									<td colspan="4" class="iptsubmit"> 
											<input type="submit" value="确定">
											<input type="reset" value="从来">										</td>
								</tr>
							</table>
			</form>
		</td>
	</tr>
</table>
<%htmlend%>
