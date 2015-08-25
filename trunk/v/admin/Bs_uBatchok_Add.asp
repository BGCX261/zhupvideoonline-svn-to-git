<!--#include file="bsconfig.asp"-->
<!--#include file="inc/md5.asp"-->
<%if Request.QueryString("no")="grow" then


set rs=server.createobject("adodb.recordset")
sqltext="select * from Bs_User where adminName='" & request("uid")(i) & "' Or realname='" & request("realname")(i) & "'"
rs.open sqltext,conn,1,1

'查找数据库，检查此管理员是否已经存在
if rs.recordcount >= 1 then
if rs("adminName")=request("uid")(i) or rs("realname")=request("realname")(i) then
Response.Redirect "Loginsb.asp?msg=此管理员帐号已经存在，请选用其它名称!"
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
<div id="sysNavPath">您的位置：新闻类别新增页</div>
<div class="jsTestAlt">进入新闻类别新增页面，要求点“增加一行“按钮时可以增加（要求最多一次增加10条），以便同时提交多条数据！”</div>
<div id="sysIptBox"><span>新闻类别新增（点增加一行，可一次新增多条）</span>
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
if(obj.value=='增加'){ 
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
objCell.innerHTML=objSourceRow.cells[11].innerHTML.replace(/增加/,'删除'); 
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
    alert("用户名不能为空！");
	document.getElementsByName("uid")[rowIndex].focus();
	return false;
  }
  if (document.getElementsByName("pwd1")[rowIndex].value=="")
  {
    alert("密码不能为空！");
	document.getElementsByName("pwd1")[rowIndex].focus();
	return false;
  }
  return true;  
}

//--> 
</script> 
 <table width="100%" border=1>
  <tr>
    <th>上传帐号</th>
    <th>真实姓名</th>
    <th>管理员密码</th>
    <th>角色身份</th>
    <th>所属院系</th>
    <th>专业班级</th>
    <th>电子邮件</th>
    <th>联系电话</th>
    <th>QQ</th>
    <th>入学年份</th>
    <th>简介</th>
    <th>操作</th>
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
          <option selected="selected" value="">请选择用户角色</option>
        </select></td>
        <td><select name="ufaculty">
          <option selected="selected" value="">请选择所属院系</option>
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
      </select></td>
    <td><input type="text" name="ubelong" size="12" maxlength="20"></td>
    <td><input type="text" name="uemail" size="12" maxlength="20"></td>
    <td><input type="text" name="uphone" size="12"></td>
    <td><input type="text" name="uqq" size="12"></td>
    <td><input type="text" name="admission" size="6"></td>
    <td><input type="text" name="uprofile" size="24"></td>
    <td><input name="add" type="button" id="add" value="增加" onClick="addLine(this)"></td>
  </tr>
 </table>
</td>
      </tr>  
    <tr>
      <td class="sysSchSmtTd"><input name="input" type="submit" value=" 保 存 " /><input name="input" type="button" value=" 返 回 " onclick="window.history.go(-1)" /></td>
    </tr>
  </table>
</form>
</div>
</div>
<%htmlend%>