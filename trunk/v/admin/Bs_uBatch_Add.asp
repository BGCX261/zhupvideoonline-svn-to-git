<!--#include file="bsconfig.asp"-->
<!--#include file="inc/md5.asp"-->
<%if Request.QueryString("no")="grow" then
for i=0 to request("tigerListSize")
set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_User where adminName='" & request("uid"&i) & "' or realname='" & request("realname"&i) & "'"
rs.open sqltext,conn,1,1

'�������ݿ⣬���˹���Ա�Ƿ��Ѿ�����
if rs.recordcount >= 1 then
if rs("adminName")=request("uid"&i) or rs("realname")=request("realname"&i) then
Response.Redirect "Loginsb.asp?msg=�˹���Ա�ʺ��Ѿ����ڣ���ѡ����������!"
response.end
rs.close
end if
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Bs_User"
rs.open sql,conn,1,3 

uid=request("uid"&i)
realname=request("realname"&i)
adminRole=request("adminRole0")
admin_Manage=request("admin_Manage0")
pwd=request("pwd"&i)
pwd=md5(pwd)
ufaculty=request("ufaculty"&i)
ubelong=request("ubelong"&i)
uemail=request("uemail"&i)
uphone=request("uphone"&i)
uqq=request("uqq"&i)
admission=request("admission"&i)
uprofile=request("uprofile"&i)

rs.addnew
rs("adminName")=uid 
rs("realname")=realname 
rs("adminRole")=adminRole
rs("admin_Manage")=admin_Manage  
rs("Password")=pwd 
rs("ufaculty")=ufaculty 
rs("ubelong")=ubelong 
rs("uemail")=uemail 
rs("uphone")=uphone 
rs("uqq")=uqq 
rs("admission")=admission 
rs("uprofile")=uprofile 

rs.update
next
rs.close
response.redirect "Bs_uBatch_Add.asp"
end if
%>
<script>

</script>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("03")%>
<div id="sysMain">
<div id="sysNavPath">����λ�ã���Ƶ�ϴ��û�����ҳ</div>
<div class="jsTestAlt">�㡰����һ�С���ťʱ��������һ�У������һ������5������</div>
 <form method="post"  action="Bs_uBatch_Add.asp?no=grow" name="myform" id="myform" onSubmit="return CheckForm()">
<input type="hidden" value="0" name="tigerListSize"> 

 <table class="sysIptTable">
 
    <tr>
      <td>
<SCRIPT language=JavaScript>
<!--
	var lineID=1;
	var linenumber=1;
	var maxLineNumber=4;
	function InsertLine()
	{
		//var i = parseInt(document.myform.tigerListSize.value,10);
		if (linenumber <= maxLineNumber){
			var outstring=	'<table  id="addLine'+lineID+'" width="100%" border="1">'+
				'<tr>'+
					'<td><input type="text" size="12" maxlength="20" name="uid'+lineID+'"></td>'+
					'<td><input type="text" size="12" maxlength="20" name="realname'+lineID+'"></td>'+
					'<td><input type="text" size="12" name="pwd'+lineID+'"></td>'+

					'<td>���ϵͳĬ��</td>'+
					'<td><select name="ufaculty'+lineID+'"><option selected="selected" value="">��ѡ������Ժϵ</option><option value="��������ϢѧԺ">��������ϢѧԺ</option><option value="���ù���ѧԺ">���ù���ѧԺ</option><option value="���ʺ����뽻��ѧԺ">���ʺ����뽻��ѧԺ</option><option value="�������ϵ">�������ϵ</option><option value="�������ϵ">�������ϵ</option><option value="����ϵ">����ϵ</option><option value="��������ϢѧԺ">��������ϢѧԺ</option><option value="˼����">˼����</option><option value="���˽���ѧԺ">���˽���ѧԺ</option><option value="����������">����������</option></select></td>'+
					'<td><input type="text" size="12" maxlength="20" name="ubelong'+lineID+'"></td>'+
					'<td><input type="text" size="12" name="uemail'+lineID+'"></td>'+
					'<td><input type="text" size="6" name="admission'+lineID+'"></td>'+
					'<td><input type="text" size="12" name="uphone'+lineID+'"></td>'+
					'<td><input type="text" size="12" name="uqq'+lineID+'"></td>'+
					'<td><input type="text" size="24" name="uprofile'+lineID+'"></td>'+
					
					'<td><INPUT value="ɾ��" type="button" onclick="SetToDel1(' + lineID + ')"></td>'+
				'</tr>'+
			'</table>';
			eval('document.getElementById("addLine").insertAdjacentHTML("AfterEnd", outstring)');
			linenumber++;
			lineID++;
		}
		document.myform.tigerListSize.value=linenumber-1;
	}

	function SetToDel1(number)
	{
		eval('document.getElementById("addLine' + number + '").style.display="none"');
		eval('document.getElementById("addLine' + number + '").outerHTML=""');
		linenumber--;
		document.myform.tigerListSize.value=linenumber-1;
	}
	
/*function CheckForm()
{

var f=document.getElementsByName("input");
alert(f.length)
for(var i=0;i<f.length;i++){
	alert(f[i])
	if(f[i].value=="")
	alert("��"+(i+1)+"��������");
	return false;
  }
  return true;  
}*/




	//-->
	</SCRIPT>
 <table width="100%" border=1>
  <tr>
    <th>�ϴ��ʺ�</th>
    <th>��ʵ����</th>
    <th>����Ա����</th>
    <th>��ɫ���</th>
    <th>����Ժϵ</th>
    <th>רҵ�༶</th>
    <th>�����ʼ�</th>
    <th>��ѧ���</th>
    <th>��ϵ�绰</th>
    <th>QQ</th>
    <th>���</th>
    <th>����</th>
  </tr>
  <tr>
    <td><input type="text" name="uid0" size="12" maxlength="20"></td>
    <td><input type="text" name="realname0" size="12" maxlength="20"></td>
    <td><input type="text" name="pwd0" size="12">
    
    </td>
    <td>
<%                
dim rs1,sql1,sel1           
 set rs1=server.createobject("adodb.recordset")           
 sql1="select * from Bs_Role where id=1"           
 rs1.open sql1,conn,1,1           
	do while not rs1.eof  %>                   
		   <input type="hidden" name="admin_Manage0" value='<%=CStr(rs1("rolePower"))%>'>
		   <input type="text" name="adminRole0" value='<%=CStr(rs1("roleName"))%>' readonly="readonly" size="12">
   <% rs1.movenext           
    loop           
	rs1.close
set rs1=nothing           
%>
    </td>
    <td>
      <select name="ufaculty0">
          <option selected="selected" value="">��ѡ������Ժϵ</option>
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
    <td><input type="text" name="ubelong0" size="12" maxlength="20"></td>
    <td><input type="text" name="uemail0" size="12"></td>
    <td><input type="text" name="admission0" size="6"></td>
    <td><input type="text" name="uphone0" size="12"></td>
    <td><input type="text" name="uqq0" size="12"></td>
    <td><input type="text" name="uprofile0" size="24"></td>
    <td></td>
  </tr>
 </table>
  <div id="addLine">
  </div>
</td>
      </tr>  
    <tr>
      <td class="sysSchSmtTd"> <input onclick="InsertLine()" type="button" value="����һ��" name="save"> <input name="input" type="submit" value=" �� �� " /><input name="input" type="button" value=" �� �� " onclick="window.history.go(-1)" /></td>
    </tr>
  </table>
</form>
</div>
</div>
<script>
function CheckForm()
{
	var form = document.forms['myform'];
	  inputs = form.getElementsByTagName('input');
	for(var i = 0; i < inputs.length; i++){
	  var input = inputs[i];
	  if(input.type === 'text' && (input.value === '' || input.value === null)){
	  alert(input.name + 'Ϊ��');
	   return false;
	  }
	} 
	return true;
}
	
</script>
<%htmlend%>