<!--#include file="bsconfig.asp"-->

<%
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_User"
rs.open sqltext,conn,1,1
%>
<!-- #include file="Inc/Head.asp" -->
<script language="javascript">
function confirmdel(id){
if (confirm("真的要删除此管理员帐号?"))
window.location.href="Bs_Admin_Del.asp?id="+id+"  "   }
</script>
<SCRIPT language=javascript id=clientEventHandlersJS>
<!--


function form1_onsubmit(){
	if (document.FORM1.pwd10.value!=document.FORM1.pwd2.value){
		alert ("请确认您的密码。");
		document.FORM1.pwd10.value='';
		document.FORM1.pwd2.value='';
		document.FORM1.pwd10.focus();
		return false;
	}
   if (document.FORM1.ufaculty.value=="")
	  {
		alert("请选择所属院系部！");
		document.FORM1.ufaculty.focus();
		return false;
	  }
}


//-->
</SCRIPT>
<%call checkmanage("01")%>
<table align="center" class="a2">
	<tr>
		<td class="a1">添加管理员</td>
	</tr>
	<tr class="a4">
		<td>
      <FORM language=javascript name=FORM1  onsubmit="return form1_onsubmit()" action="Bs_Admin_Add.asp" method=post>
			<table width="100%">
				<tr> 
					<td class="iptlab">管理员帐号：</td>
					<td class="ipt"><input type="text" name="uid" maxlength="20"></td>
					<td class="iptlab">真实姓名：</td>
					<td class="ipt"><input type="text" name="realname" maxlength="20"></td>
				</tr>
				<tr> 
					<td class="iptlab"> 管理员密码：</td>
					<td class="ipt"><input type="text" name="pwd10"></td>
					<td class="iptlab">密码确认：</td>
					<td class="ipt"><input type="text" name="pwd2"></td>
				</tr>
                <tr> 
					<td class="iptlab">专业班级或教研室：</td>
					<td class="ipt"><input type="text" name="ubelong"  maxlength="20"></td>
					<td class="iptlab">电子邮件：</td>
					<td class="ipt"><input type="text" name="uemail" maxlength="20"></td>
				</tr>
				<tr> 
					<td class="iptlab"> 联系电话：</td>
					<td class="ipt"><input type="text" name="uphone" ></td>
					<td class="iptlab">QQ：</td>
					<td class="ipt"><input type="text" name="uqq" ></td>
				</tr>
                <tr> 
                    <td class="iptlab">身份角色：</td>
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
          <option selected="selected" value="">请选择用户角色</option>
        </select>
                    </td>
                    <td class="iptlab"> 入学年份：</td>
					<td class="ipt"><input type="text" name="admission"></td>
				</tr>
                <tr> 
					<td class="iptlab"> 简 介：</td>
					<td class="ipt"><input type="text" name="uprofile" size="36"></td>
                    <td class="iptlab"> 所在院系部：</td>
					<td class="ipt">
                    <select name="ufaculty">
                      <option selected="selected" value="">请选择所在院系部</option>
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
					<td colspan="4" class="iptsubmit"><INPUT type=submit value='确认添加' name=Submit2></td>
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
	<tr class="a4">
		<td>
			<table width="100%">
				<tr class="opttitle"> 
					<td>管理员帐号</td>
					<td>真实姓名</td>
                    <td>用户身份</td>
					<td>管理员密码</td>
					<td>操作</td>
					<td>删除</td>
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
		</td>
	</tr>
</table>
<%htmlend%>
