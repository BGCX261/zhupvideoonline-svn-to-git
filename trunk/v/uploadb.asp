<!--#include file="Inc/config.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
BODY{
BACKGROUND-COLOR: #f5feed;
font-size:9pt
}
.tx1 { height: 20px;font-size: 9pt; border: 1px solid; border-color: #000000; color: #0000FF}
-->
</style>

<SCRIPT language=javascript>
function check_file() 
{
  var strFileName=form1.FileName.value;
  var range = parent.myform.honor;
  if (strFileName=="")
  {
    alert("请选择要上传的文件");
    return false;
  }
}
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0">
<%
if EnableUploadFile="Yes" and (session("adminName")<>"" or session("UserName")<>"") then
%>
<form action="upfileb.asp" method="post" name="form1" onSubmit="return check_file()" enctype="multipart/form-data">
  <input name="FileName" type="FILE" class="tx1" size="20">
  <input type="submit" name="Submit" value="上传" style="border:1px double rgb(88,88,88);font:9pt">
</form>
<%
end if
%>
</body>
</html>
