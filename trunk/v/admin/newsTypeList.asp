<!--#include file="../conn.asp"-->
<!--#include file="../commonFunction.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../../public/css/adminStyle.css" type="text/css" rel="stylesheet" />
<script language="javascript" src="../../public/js/ctrl_checkbox.js"></script>
<script>
function add(url)
{
     window.location=url;
}
function mdy()
{
	if (getCheckedNum(document.forms[1].checkme) > 1)
	{
		alert("每次只能修改一条记录");
		return;
	}
	else if (getCheckedNum(document.forms[1].checkme) == 0)
	{
		alert("请先选择一条记录");
		
		return;
	}

        document.forms[1].target="_self";
		document.forms[1].action="newsTypeModify.asp";
	    document.forms[1].submit();
}
function del(strId){
if (getCheckedNum(document.forms[1].checkme) == 0)
	{
		alert("请先选择一条或者多条记录");
		return;
	}
        var res = confirm("是否真的要删除？");
        if(res == true){
	    document.forms[1].target="_self";
		document.forms[1].action="newsTypeDel.asp";
	    document.forms[1].submit();
        }
}

</script>
<%
id=replace(request("checkme")," ","")
%>
<title>新闻类别管理页面</title>
</head>

<body>
<div id="sysMain">
<div id="sysNavPath">您的位置：新闻类别管理列表页</div>
<div class="jsTestAlt">进入新闻管理列表页面，要求修改和删除通过多选框的选择来操作，修改js判断只能选择单条，删除可以是多条。没有选择则提示“请选择后再操作！”</div>
<div id="doTool"><input name="" type="button" value=" 增 加 "  onclick="add('newsTypeAdd.asp')" /><input name="" type="button" value=" 修 改 "  onclick="mdy()" /><input name="" type="button" value=" 删 除 "  onClick="del();" /></div>
<div id="sysSch"><span>新闻类别查询</span>
<form method="get" name="myform" action="newsTypeList.asp">
  <table class="sysSchTable">
    <tr>
      <td>用户名</td>
      <td>
        <input type="text" name="newsTypeName" id="newsTypeName"  value="请输用户名" onfocus="if(this.value=='请输入用户名'){this.value=''};" /></td>
      <td>真实姓名</td>
      <td><input type="text" name="newsTypeName2" id="newsTypeName2"  /></td>
      <td>专业班级</td>
      <td><input type="text" name="newsTypeName3" id="newsTypeName3"   /></td>
    </tr>
        <tr>
      <td>状态</td>
      <td>
        <select name="uStatus">
          <option value="2">全部</option>
          <option value="1">启用</option>
          <option value="0">停用</option>
        </select>
      </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td colspan="6" class="sysSchSmtTd"><input name="input" type="submit" value=" 查 询 " /></td>
      </tr>
  </table>
  </form>
</div>
<div class="comnDiv">
<form>
    <%
newsTypeName=Trim(request("newsTypeName"))
	sql="select * from Bs_User where ID>0"
if newsTypeName="" or newsTypeName="请输入类别名称" then
	sql=sql
else
	sql=sql & " and news_Type_Name like '%" & newsTypeName & "%' "
end if
sql=sql & " order by ID desc"
	
		set rsSysAdmin=server.CreateObject("adodb.recordset")
		rsSysAdmin.open sql,conn,1,1
		if (rsSysAdmin.eof and rsSysAdmin.bof) then
			response.Write("无数据！！请添加！！")
		else
			'--- 1 分页初始化 -----
			pSize=10	'每页显示记录数
			call pageInit(rsSysAdmin,pSize  ,currentPage,pageSize)
    %>
  <table class="sysListTable">
    <tr class="sysListTitleTr">
      <td>全选
        <input name="checkbox" type="checkbox" id="checkall" title="全选" onClick="if(document.getElementById('checkme') != null){checkAll(checkme, checkall);}"></td>
      <td>ID</td>
      <td>用户名</td>
      <td>真实姓名</td>
      <td>专业班级</td>
      <td>电子邮件</td>
      <td>联系电话</td>
      <td>QQ</td>
      <td>状态</td>
      <td>加入时间</td>
      <td>简介</td>
      </tr>
    <tbody id="goaler">
   <%
				do while(not rsSysAdmin.eof)
					'--- 2 循环处设界 -----
					if (pageSize<=0)then exit do end if
						pageSize=pageSize-1
    %>
    <tr class="sysListSingleTr">
      <td><input type="checkbox" name="checkme" value="<%=rsSysAdmin("ID")%>" id="checkme"  onclick="onClickChk(this, checkme, checkall)"></td>
      <td><%=rsSysAdmin("ID")%></td>
      <td><%=rsSysAdmin("adminName")%></td>
      <td><%=rsSysAdmin("realname")%></td>
      <td><%=rsSysAdmin("ubelong")%></td>
      <td><%=rsSysAdmin("uemail")%></td>
      <td><%=rsSysAdmin("uphone")%></td>
      <td><%=rsSysAdmin("uqq")%></td>
      <td><%=rsSysAdmin("uStatus")%></td>
      <td><%=rsSysAdmin("uTime")%></td>
      <td><%=rsSysAdmin("uprofile")%></td>
      </tr>
    <%
				 	rsSysAdmin.moveNext
				 loop
    %>
    </tbody>
  </table>
  <%end if%>
  </form>
</div>
<div id="turnPage">    
    <%
		'--- 3 分页显示 -----
        call pagination(rsSysAdmin	,currentPage)
    %></div>
</div>
</body>
</html>
<script>
//列表表格隔行变色
var TbRow = document.getElementById("goaler");
if (TbRow != null){
   for (var i=0;i<TbRow.rows.length ;i++ )
     {
        if (TbRow.rows[i].rowIndex%2==1){
            TbRow.rows[i].className="sysListSingleTr";
        }
        else{
            TbRow.rows[i].className="sysListDoubleTr";
        }
     }
}
</script>
<%
call closeConnDB()
%>