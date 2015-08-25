<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
<!--
body{BACKGROUND: #f5feed;font-size:12px; margin:0;}
.tx1 { height: 20px;font-size:12px; border: 1px solid #000; color: #00F;}
-->
</style>
<SCRIPT language=javascript>
function check_file() 
{
  var strFileName=form1.FileName.value;
  var range = parent.cgfans.minPic;
  if (strFileName=="")
  {
    alert("请选择要上传的文件");
    return false;
  }
}
</SCRIPT>
</head>
<body>
<%
if ( session("UserName")<>"") then
%>
<form action="upfilea.asp" method="post" name="form1" onSubmit="return check_file()" enctype="multipart/form-data">
  <input name="FileName" type="file" class="tx1" size="20">
  <input type="submit" name="submit" value=" 上传缩略图 " style="border:1px double rgb(88,88,88);font:12px;">
</form>
<%
end if
%>
</body>
</html>
