<!--#include file="../conn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../../public/css/adminStyle.css" type="text/css" rel="stylesheet" />
<title>新闻类别修改页面</title>
<%
dim OldNewsTypeId
if Request.QueryString("no")="grow" then
id=request("newsTypeId")
newsTypeName=request("newsTypeName")
OldNewsTypeName=request("OldNewsTypeName")
newsTypeProfile=Trim(Request("newsTypeProfile"))



If newsTypeName="" Then
response.write "SORRY <br>"
response.write "请输入新闻类别名称"
response.end
end if
If newsTypeProfile="" Then
response.write "SORRY <br>"
response.write "新闻类别简介不能为空"
response.end
end if


Set rsSysAdmin = Server.CreateObject("ADODB.Recordset")
sql="select * from Bs_User where news_Type_Id="&id
rsSysAdmin.open sql,conn,1,3

rsSysAdmin("news_Type_Name")=newsTypeName
rsSysAdmin("news_Type_Profile")=newsTypeProfile
rsSysAdmin("news_Type_Mdy_Date")=now()
rsSysAdmin.update
rsSysAdmin.close
if newsTypeName<>OldNewsTypeName then
	conn.execute "Update Bs_User set news_Type_Name='" & newsTypeName & "' where news_Type_Name='" & OldNewsTypeName & "'"
end if	
response.redirect "newsTypeList.asp"
end if
%>
<%
id=request("checkme")
Set rsSysAdmin = Server.CreateObject("ADODB.Recordset")
rsSysAdmin.Open "select * from Bs_User where news_Type_Id="&id, conn,3,3
%>
<script>
function CheckForm()
{
  if (document.myform.newsTypeName.value=="")
  {
    alert("用户名不能为空！");
	document.myform.newsTypeName.focus();
	return false;
  }
  return true;  
}
</script>
</head>

<body>
<div id="sysMain">
<div id="sysNavPath">您的位置：新闻类别修改页</div>
<div class="jsTestAlt">进入新闻类别修改页面，修改请做好各个输入框的js校验！”</div>
<div id="sysIptBox"><span>新闻类别修改</span>
<form method="post" action="newsTypeModify.asp?no=grow" name="myform" onSubmit="return CheckForm()">
<input type="hidden" name="newsTypeId" value=<%=id%>>
<input name="OldNewsTypeName" type="hidden" id="OldNewsTypeName" value="<%=rsSysAdmin("news_Type_Name")%>"> 
  <table class="sysIptTable">
    <tr>
      <td>类别名称</td>
      <td>
        <input name="newsTypeName" type="text" id="newsTypeName" value="<%=rsSysAdmin("news_Type_Name")%>" /></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td >简介（不超过100字）</td>
      <td colspan="5" >
        <textarea name="newsTypeProfile" id="newsTypeProfile" cols="45" rows="5"><%=rsSysAdmin("news_Type_Profile")%></textarea></td>
    </tr>
    <tr>
      <td colspan="6" class="sysSchSmtTd"><input name="input" type="submit" value=" 保 存 " /> <input name="input" type="button" value=" 返 回 " onclick="window.history.go(-1)" /></td>
      </tr>
  </table>
    </form>
</div>
<div class="comnDiv"></div>
</div>
</body>
</html>
<%
call closeConnDB()
%>