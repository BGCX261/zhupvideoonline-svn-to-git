<!--#include file="bsconfig.asp"-->

<%
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_User"
rs.open sqltext,conn,1,1
%>
<!-- #include file="Inc/Head.asp" -->
<script language="javascript">
function confirmdel(id){
if (confirm("���Ҫɾ���˹���Ա�ʺ�?"))
window.location.href="Bs_Admin_Del.asp?id="+id+"  "   }
</script>
<SCRIPT language=javascript id=clientEventHandlersJS>
<!--


function form1_onsubmit(){
	if (document.FORM1.pwd10.value!=document.FORM1.pwd2.value){
		alert ("��ȷ���������롣");
		document.FORM1.pwd10.value='';
		document.FORM1.pwd2.value='';
		document.FORM1.pwd10.focus();
		return false;
	}
   if (document.FORM1.ufaculty.value=="")
	  {
		alert("��ѡ������Ժϵ����");
		document.FORM1.ufaculty.focus();
		return false;
	  }
}


//-->
</SCRIPT>
<%call checkmanage("01")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">��ӹ���Ա</td>
	</tr>
	<tr class="a4">
		<td>
      <FORM language=javascript name=FORM1  onsubmit="return form1_onsubmit()" action="Bs_Admin_Add.asp" method=post>
			<table width="100%">
				<tr> 
					<td class="iptlab">����Ա�ʺţ�</td>
					<td class="ipt"><input type="text" name="uid" maxlength="20"></td>
					<td class="iptlab">��ʵ������</td>
					<td class="ipt"><input type="text" name="realname" maxlength="20"></td>
				</tr>
				<tr> 
					<td class="iptlab"> ����Ա���룺</td>
					<td class="ipt"><input type="text" name="pwd10"></td>
					<td class="iptlab">����ȷ�ϣ�</td>
					<td class="ipt"><input type="text" name="pwd2"></td>
				</tr>
                <tr> 
					<td class="iptlab">רҵ�༶������ң�</td>
					<td class="ipt"><input type="text" name="ubelong"  maxlength="20"></td>
					<td class="iptlab">�����ʼ���</td>
					<td class="ipt"><input type="text" name="uemail" maxlength="20"></td>
				</tr>
				<tr> 
					<td class="iptlab"> ��ϵ�绰��</td>
					<td class="ipt"><input type="text" name="uphone" ></td>
					<td class="iptlab">QQ��</td>
					<td class="ipt"><input type="text" name="uqq" ></td>
				</tr>
                <tr> 
                    <td class="iptlab">��ݽ�ɫ��</td>
					<td class="ipt">
                    <select name="adminRole">
          <%                
dim rs1,sql1,sel1           
 set rs1=server.createobject("adodb.recordset")           
  sql1="select * from Bs_Role"           
 rs1.open sql1,conn,1,1           
			  do while not rs1.eof           
                                sel="selected"               
		             response.write "<option " & sel & " value=" & rs1("rolePower") & ">" & rs1("roleName") & "</option>"+chr(13)+chr(10)           
		             rs1.movenext           
    		          loop           
			rs1.close
			set rs1=nothing           
			      %>
          <option selected="selected" value="">��ѡ���û���ɫ</option>
        </select>
                    </td>
                    <td class="iptlab"> ��ѧ��ݣ�</td>
					<td class="ipt"><input type="text" name="admission"></td>
				</tr>
                <tr> 
					<td class="iptlab"> �� �飺</td>
					<td class="ipt"><input type="text" name="uprofile" size="36"></td>
                    <td class="iptlab"> ����Ժϵ����</td>
					<td class="ipt">
                    <select name="ufaculty">
                      <option selected="selected" value="">��ѡ������Ժϵ��</option>
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
					<td colspan="4" class="iptsubmit"><INPUT type=submit value='ȷ�����' name=Submit2></td>
				</tr>
			</table>
      </form>
		</td>
	</tr>
</table>
<BR>
<BR>
<table align="center" class="a2">
	<tr>
		<td class="a1">����Ա�ʺŹ���</td>
	</tr>
	<tr class="a4">
		<td>
			<table width="100%">
				<tr class="opttitle"> 
					<td>����Ա�ʺ�</td>
					<td>��ʵ����</td>
                    <td>�û����</td>
					<td>����Ա����</td>
					<td>����</td>
					<td>ɾ��</td>
				</tr>
<%
if not rs.eof then
do while not rs.eof
%>
				<tr class="opt"> 
					<td><%=rs("adminName")%></td>
					<td><%=rs("realname")%></td>
                    <td><%=rs("adminRole")%></td>
					<td><%=rs("PassWord")%></td>
					<td><%response.write "<a href='Bs_Admin_Pass_edit.asp?ID="&rs("Id")&"'  >�޸�����</a>"%></td>
					<td><%response.write "<a href='javascript:confirmdel(" & rs("Id") & ")'>ɾ��</a>"%></td>
				</tr>
<%
rs.movenext
loop
end if
%>
<%
rs.close
conn.close
%>
			</table>
		</td>
	</tr>
</table>
<%htmlend%>
