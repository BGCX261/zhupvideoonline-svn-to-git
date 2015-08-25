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
if (confirm("真的要删除此管理员帐号?"))
window.location.href="Bs_Admin_Del.asp?id="+id+"  "   }
</script>
<SCRIPT language=javascript id=clientEventHandlersJS>
<!--


function form1_onsubmit()
{

if(document.FORM1.pwd1.value==""){
   alert("密码不能为空！");
   document.FORM1.pwd1.focus();
   return false;
}

if(document.FORM1.realname.value==""){
   alert("真实姓名不能为空！");
   document.FORM1.realname.focus();
   return false;
}

if (document.FORM1.pwd1.value!=document.FORM1.pwd2.value)
{
alert ("请确认您的密码。");
document.FORM1.pwd1.value='';
document.FORM1.pwd2.value='';
document.FORM1.pwd1.focus();
return false;
}

if(document.FORM1.uid.value==""){
   alert("用户名不能为空！");
   document.FORM1.uid.focus();
   return false;
}

if(document.FORM1.purview.value==""){
   alert("请选择用户身份！");
   document.FORM1.purview.focus();
   return false;
}

}


//-->
</SCRIPT>
<table align="center" class="a2">
	<tr>
		<td class="a1">添加管理员</td>
	</tr>
	<tr class="a4">
		<td>
      <FORM language=javascript name=FORM1  onsubmit="return form1_onsubmit()" action="Bs_Admin_Add.asp" method=post>
            <input   type=Hidden   name=identity   value=""> 
			<table width="100%">
				<tr> 
					<td class="iptlab">管理员帐号：</td>
					<td class="ipt"><input type="text" name="uid" size="16" maxlength="20"></td>
				</tr>
				<tr> 
					<td class="iptlab">真实姓名：</td>
					<td class="ipt"><input type="text" name="realname" size="16" maxlength="20"></td>
				</tr>
				<tr> 
					<td class="iptlab"> 管理员密码：</td>
					<td class="ipt"><input type="text" name="pwd1" size="16"></td>
				</tr>
				<tr> 
					<td class="iptlab">密码确认：</td>
					<td class="ipt"><input type="text" name="pwd2" size="16"></td>
				</tr>
                <tr> 
					<td class="iptlab">身份：</td>
					<td class="ipt">
                    <select name="purview" id="purview" onChange="identity.value=this.options[this.selectedIndex].text;">
									      <option value="" selected>请选择用户身份</option>
                                          <option value="0">老师</option>
                                          <option value="1">辅导员</option>
                                          <option value="2">学生</option>
                                          <option value="3">其他</option>
								        </select>
                    </td>
				</tr>
				<tr> 
					<td colspan="2" class="iptsubmit"><INPUT type=submit value='确认添加' name=Submit2></td>
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
		<td class="a1">管理员帐号管理</td>
	</tr>
    <tr>
		<td>添加完管理员帐户后请到“权限管理”为管理员赋予相应的权限。</td>
	</tr>
	<tr class="a4">
		<td>
			<FORM language=javascript action="Bs_Admin_Add.asp" method=post>
			<table width="100%">
				<tr class="opttitle"> 
					<td>管理员帐号</td>
					<td>真实姓名</td>
					<td>身份</td>
					<td>操作</td>
					<td>删除</td>
				</tr>
<%
if not rs.eof then
do while not rs.eof
%>
				<tr class="opt"> 
					<td><%=rs("UserName")%></td>
					<td><%=rs("realname")%></td>
					<td><%=rs("identity")%></td>
					<td><%response.write "<a href='Bs_Admin_Pass_edit.asp?ID="&rs("Id")&"'  >修改密码</a>"%></td>
					<td><%response.write "<a href='javascript:confirmdel(" & rs("Id") & ")'>删除</a>"%></td>
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
