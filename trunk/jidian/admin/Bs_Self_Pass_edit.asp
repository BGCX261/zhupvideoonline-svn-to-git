<!--#include file="bsconfig.asp"-->

<%
id=request("id")
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_User where Id=" & id
rs.open sqltext,conn,1,1
%>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("11")%>
<SCRIPT language=javascript>
<!--
function form1_onsubmit(){
var pwd1=document.form1.pwd1.value;
var pwd2=document.form1.pwd2.value;
if(pwd1==""){
   alert("���벻��Ϊ�գ�");
   document.form1.pwd1.focus();
   return false;
}

if ( pwd1.length < 4 ) {
  		alert("���벻������4�ֽڣ�");
		document.form1.pwd1.focus();
		return false;
}

if(pwd1!=pwd2){
   alert("������������벻һ�£����������롣");
   document.form1.pwd1.focus();
   return false;
}

  } 	
//-->
</SCRIPT>
<table align="center" class="a2">
	<tr>
		<td class="a1">����Ա�����޸�</td>
	</tr>
	<tr class="a4">
		<td>
			<FORM  name=form1  onsubmit="return form1_onsubmit()" action=Bs_Admin_Pass_Save.asp?uid=<%=rs("Id")%> method=post>
			<INPUT type=hidden value=<%=rs("UserName")%> name=uid>
				<table width="100%">
					<tr> 
						<td class="iptlab">����Ա�ʺţ�</td>
						<td class="ipt"><%=rs("UserName")%></td>
					</tr>
					<tr> 
						<td class="iptlab">�����룺</td>
						<td class="ipt"><input name="pwd1" type="text" size="16" maxlength="20"></td>
					</tr>
					<tr> 
						<td class="iptlab">����ȷ�ϣ�</td>
						<td class="ipt"><input name="pwd2" type="text" size="16" maxlength="20"></td>
					</tr>
					<tr> 
						<td colspan="2" class="iptsubmit"><INPUT  type=submit value='ȷ���޸�' name=Submit2></td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
</table>
<%htmlend%>
