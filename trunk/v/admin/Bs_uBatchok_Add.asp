<!--#include file="bsconfig.asp"-->
<!--#include file="inc/md5.asp"-->
<%if Request.QueryString("no")="grow" then


set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_User where adminName='" & request("uid")(i) & "' Or realname='" & request("realname")(i) & "'"
rs.open sqltext,conn,1,1

'�������ݿ⣬���˹���Ա�Ƿ��Ѿ�����
if rs.recordcount >= 1 then
if rs("adminName")=request("uid")(i) or rs("realname")=request("realname")(i) then
Response.Redirect "Loginsb.asp?msg=�˹���Ա�ʺ��Ѿ����ڣ���ѡ����������!"
response.end
rs.close
end if
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Bs_User"
rs.open sql,conn,1,3 

for  i=1  to  request("uid").count

uid=request("uid")(i)
realname=request("realname")(i)
adminRole=request("adminRole")(i)
admin_Manage=request("admin_Manage")(i)
pwd1=request("pwd1")(i)
pwd1=md5(pwd1)
ufaculty=request("ufaculty")(i)
ubelong=request("ubelong")(i)
uemail=request("uemail")(i)
uphone=request("uphone")(i)
uqq=request("uqq")(i)
admission=request("admission")(i)
uprofile=request("uprofile")(i)

rs.addnew
rs("adminName")=uid 
rs("realname")=realname 
rs("adminRole")=adminRole
rs("admin_Manage")=admin_Manage  
rs("Password")=pwd1 
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
response.redirect "Bs_uBatchok_Add.asp"
end if
%>
<!-- #include file="Inc/Head.asp" -->
<%call checkmanage("03")%>
<div id="sysMain">
<div id="sysNavPath">����λ�ã������������ҳ</div>
<div class="jsTestAlt">���������������ҳ�棬Ҫ��㡰����һ�С���ťʱ�������ӣ�Ҫ�����һ������10�������Ա�ͬʱ�ύ�������ݣ���</div>
<div id="sysIptBox"><span>�������������������һ�У���һ������������</span>
 <form method="post"  action="Bs_uBatchok_Add.asp?no=grow" name="myform" id="myform" onSubmit="return CheckForm()">
 <table class="sysIptTable">
 
    <tr>
      <td>
<script language="JavaScript"> 
<!-- 

var rowIndex=0; 
function addLine(obj){ 
var objSourceRow=obj.parentNode.parentNode; 
var objTable=obj.parentNode.parentNode.parentNode.parentNode; 
if(obj.value=='����'){ 
rowIndex++; 
var objRow=objTable.insertRow(rowIndex); 
var objCell; 

objCell=objRow.insertCell(0); 
objCell.innerHTML=objSourceRow.cells[0].innerHTML; 

objCell=objRow.insertCell(1); 
objCell.innerHTML=objSourceRow.cells[1].innerHTML; 

objCell=objRow.insertCell(2); 
objCell.innerHTML=objSourceRow.cells[2].innerHTML; 

objCell=objRow.insertCell(3); 
objCell.innerHTML=objSourceRow.cells[3].innerHTML; 

objCell=objRow.insertCell(4); 
objCell.innerHTML=objSourceRow.cells[4].innerHTML; 

objCell=objRow.insertCell(5); 
objCell.innerHTML=objSourceRow.cells[5].innerHTML; 

objCell=objRow.insertCell(6); 
objCell.innerHTML=objSourceRow.cells[6].innerHTML; 

objCell=objRow.insertCell(7); 
objCell.innerHTML=objSourceRow.cells[7].innerHTML; 

objCell=objRow.insertCell(8); 
objCell.innerHTML=objSourceRow.cells[8].innerHTML; 

objCell=objRow.insertCell(9); 
objCell.innerHTML=objSourceRow.cells[9].innerHTML; 

objCell=objRow.insertCell(10); 
objCell.innerHTML=objSourceRow.cells[10].innerHTML; 

objCell=objRow.insertCell(11); 
objCell.innerHTML=objSourceRow.cells[11].innerHTML.replace(/����/,'ɾ��'); 
} 
else{ 
objTable.lastChild.removeChild(objSourceRow); 
rowIndex--; 
} 
}

function slt(){ 
document.getElementsByName('adminRole')[rowIndex].value=document.getElementsByName("admin_Manage")[rowIndex].options[document.getElementsByName("admin_Manage")[rowIndex].selectedIndex].text;
} 

function removeLine(){ 

} 

function CheckForm()
{
  if (document.getElementsByName("uid")[rowIndex].value=="")
  {
    alert("�û�������Ϊ�գ�");
	document.getElementsByName("uid")[rowIndex].focus();
	return false;
  }
  if (document.getElementsByName("pwd1")[rowIndex].value=="")
  {
    alert("���벻��Ϊ�գ�");
	document.getElementsByName("pwd1")[rowIndex].focus();
	return false;
  }
  return true;  
}

//--> 
</script> 
 <table width="100%" border=1>
  <tr>
    <th>�ϴ��ʺ�</th>
    <th>��ʵ����</th>
    <th>����Ա����</th>
    <th>��ɫ���</th>
    <th>����Ժϵ</th>
    <th>רҵ�༶</th>
    <th>�����ʼ�</th>
    <th>��ϵ�绰</th>
    <th>QQ</th>
    <th>��ѧ���</th>
    <th>���</th>
    <th>����</th>
  </tr>
  <tr>
    <td><input type="text" name="uid" size="12" maxlength="20"></td>
    <td><input type="text" name="realname" size="12" maxlength="20"></td>
    <td><input type="text" name="pwd1" size="12" />      <input type="hidden" name="adminRole" id="adminRole" size="16"></td>
    <td><select name="admin_Manage" id="admin_Manage" onchange="slt()">
          <%                
dim rs1,sql1,sel1           
 set rs1=server.createobject("adodb.recordset")           
  sql1="select * from Bs_Role"           
 rs1.open sql1,conn,1,1           
			  do while not rs1.eof           
                                sel="selected"               
		             response.write "<option " & sel & " value='"+CStr(rs1("rolePower"))+"'>"+rs1("roleName")+"</option>"+chr(13)+chr(10)           
		             rs1.movenext           
    		          loop           
			rs1.close
			set rs1=nothing           
			      %>
          <option selected="selected" value="">��ѡ���û���ɫ</option>
        </select></td>
        <td><select name="ufaculty">
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
      </select></td>
    <td><input type="text" name="ubelong" size="12" maxlength="20"></td>
    <td><input type="text" name="uemail" size="12" maxlength="20"></td>
    <td><input type="text" name="uphone" size="12"></td>
    <td><input type="text" name="uqq" size="12"></td>
    <td><input type="text" name="admission" size="6"></td>
    <td><input type="text" name="uprofile" size="24"></td>
    <td><input name="add" type="button" id="add" value="����" onClick="addLine(this)"></td>
  </tr>
 </table>
</td>
      </tr>  
    <tr>
      <td class="sysSchSmtTd"><input name="input" type="submit" value=" �� �� " /><input name="input" type="button" value=" �� �� " onclick="window.history.go(-1)" /></td>
    </tr>
  </table>
</form>
</div>
</div>
<%htmlend%>