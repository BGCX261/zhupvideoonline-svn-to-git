<!--#include file="bsconfig.asp"-->

<%
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_User"
rs.open sqltext,conn,1,1
%>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("01")%>
<script language="javascript">
function confirmdel(id){
if (confirm("���Ҫɾ���˹���Ա�ʺ�?"))
window.location.href="Bs_Admin_Del.asp?id="+id+"  "   }
</script>
<SCRIPT language=javascript id=clientEventHandlersJS>
<!--


function form1_onsubmit()
{

if(document.FORM1.pwd1.value==""){
   alert("���벻��Ϊ�գ�");
   document.FORM1.pwd1.focus();
   return false;
}

if(document.FORM1.realname.value==""){
   alert("��ʵ��������Ϊ�գ�");
   document.FORM1.realname.focus();
   return false;
}

if (document.FORM1.pwd1.value!=document.FORM1.pwd2.value)
{
alert ("��ȷ���������롣");
document.FORM1.pwd1.value='';
document.FORM1.pwd2.value='';
document.FORM1.pwd1.focus();
return false;
}

if(document.FORM1.uid.value==""){
   alert("�û�������Ϊ�գ�");
   document.FORM1.uid.focus();
   return false;
}

if(document.FORM1.purview.value==""){
   alert("��ѡ���û���ݣ�");
   document.FORM1.purview.focus();
   return false;
}

}


//-->
</SCRIPT>
<table align="center" class="a2">
	<tr>
		<td class="a1">��ӹ���Ա</td>
	</tr>
	<tr class="a4">
		<td>
      <FORM language=javascript name=FORM1  onsubmit="return form1_onsubmit()" action="Bs_Admin_Add.asp" method=post>
            <input   type=Hidden   name=identity   value=""> 
			<table width="100%">
				<tr> 
					<td class="iptlab">����Ա�ʺţ�</td>
					<td class="ipt"><input type="text" name="uid" size="16" maxlength="20"></td>
				</tr>
				<tr> 
					<td class="iptlab">��ʵ������</td>
					<td class="ipt"><input type="text" name="realname" size="16" maxlength="20"></td>
				</tr>
				<tr> 
					<td class="iptlab"> ����Ա���룺</td>
					<td class="ipt"><input type="text" name="pwd1" size="16"></td>
				</tr>
				<tr> 
					<td class="iptlab">����ȷ�ϣ�</td>
					<td class="ipt"><input type="text" name="pwd2" size="16"></td>
				</tr>
                <tr> 
					<td class="iptlab">��ݣ�</td>
					<td class="ipt">
                    <select name="purview" id="purview" onChange="identity.value=this.options[this.selectedIndex].text;">
									      <option value="" selected>��ѡ���û����</option>
                                          <option value="0">��ʦ</option>
                                          <option value="1">����Ա</option>
                                          <option value="2">ѧ��</option>
                                          <option value="3">����</option>
								        </select>
                    </td>
				</tr>
				<tr> 
					<td colspan="2" class="iptsubmit"><INPUT type=submit value='ȷ�����' name=Submit2></td>
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
    <tr>
		<td>��������Ա�ʻ����뵽��Ȩ�޹���Ϊ����Ա������Ӧ��Ȩ�ޡ�</td>
	</tr>
	<tr class="a4">
		<td>
			<FORM language=javascript action="Bs_Admin_Add.asp" method=post>
			<table width="100%">
				<tr class="opttitle"> 
					<td>����Ա�ʺ�</td>
					<td>��ʵ����</td>
					<td>���</td>
					<td>����</td>
					<td>ɾ��</td>
				</tr>
<%
if not rs.eof then
do while not rs.eof
%>
				<tr class="opt"> 
					<td><%=rs("UserName")%></td>
					<td><%=rs("realname")%></td>
					<td><%=rs("identity")%></td>
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
			</form>
		</td>
	</tr>
</table>

<%htmlend%>
