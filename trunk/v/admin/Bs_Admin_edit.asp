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
response.write "��������ʵ����"
response.end
end if
If uprofile="" Then
response.write "SORRY <br>"
response.write "���ݲ���Ϊ��"
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
    alert("���ݲ���Ϊ�գ�");
	editor.HtmlEdit.focus();
	return false;
  }
  return true;  
}
</script>
<!-- #include file="Inc/Head.asp" -->
<table align="center" class="a2">
	<tr>
		<td class="a1">�� �� �� �� Ա �� Ϣ</td>
	</tr>
	<tr class="a4">
		<td>
			<form method="post" action="Bs_Admin_edit.asp?no=eshop" name="myform" onSubmit="return CheckForm()">
			<input type=hidden name=id value=<%=id%>>
			<table width="100%">
								<tr> 
									<td class="iptlab">�û���</td>
									<td class="ipt"><input type="text" name="adminName" value="<%=rs("adminName")%>" readonly="readonly"></td>
                                    <td class="iptlab">��ʵ����</td>
									<td class="ipt"><input name="realname" type="text" value="<%=rs("realname")%>" readonly="readonly"></td>
								</tr>
                                <tr> 
									<td class="iptlab">��ɫ���</td>
									<td class="ipt"><input type="text" name="adminRole" value="<%=rs("adminRole")%>" readonly="readonly" /></td>
                                    <td class="iptlab">רҵ�༶</td>
									<td class="ipt"><input type="text" name="ubelong" value="<%=rs("ubelong")%>"></td>
								</tr>
                                <tr> 
									<td class="iptlab">�����ʼ�</td>
									<td class="ipt"><input type="text" name="uemail" value="<%=rs("uemail")%>"></td>
                                    <td class="iptlab">��ϵ�绰</td>
									<td class="ipt"><input type="text" name="uphone" value="<%=rs("uphone")%>"></td>
								</tr>
                                <tr> 
									<td class="iptlab">QQ</td>
									<td class="ipt"><input type="text" name="uqq" value="<%=rs("uqq")%>"></td>
                                    <td class="iptlab">��ѧ���</td>
									<td class="ipt"><input name="admission" type="text" value="<%=rs("admission")%>" readonly="readonly"></td>
								</tr>
								<tr> 
									<td class="iptlab">����</td>
									<td class="ipt">
                                    <input name="uprofile" type="text" size="36" id="uprofile" value="<%=rs("uprofile")%>"></td>
                                    <td class="iptlab"> ����Ժϵ����</td>
                                    <td class="ipt">
                                    <select name="ufaculty">
                                      <option selected="selected" value="<%=rs("ufaculty")%>"><%=rs("ufaculty")%></option>
                                      <option value="��������ϢѧԺ">��������ϢѧԺ</option>
                                      <option value="���ù���ѧԺ">���ù���ѧԺ</option>
                                      <option value="���ʺ����뽻��ѧԺ">���ʺ����뽻��ѧԺ</option>
                                      <option value="�������ϵ">�������ϵ</option>
                                      <option value="�������ϵ">�������ϵ</option>
                                      <option value="����ϵ">����ϵ</option>
                                      <option value="��������ϢѧԺ">��������ϢѧԺ</option>
                                      <option value="˼����">˼����</option>
                                      <option value="���˽���ѧԺ">���˽���ѧԺ</option>
                                      <option value="����������">����������</option>
                                    </select>
                                    </td>
                                </tr>
								<tr> 
									<td colspan="4" class="iptsubmit"> 
											<input type="submit" value="ȷ��">
											<input type="reset" value="����">										</td>
								</tr>
							</table>
			</form>
		</td>
	</tr>
</table>
<%htmlend%>
