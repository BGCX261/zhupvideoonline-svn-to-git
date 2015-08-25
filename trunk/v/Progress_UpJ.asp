<!--#include file="Inc/config.asp"-->
<html>
<head>
<title>带进度条</title>
<style type="text/css">
<!--
BODY{
BACKGROUND-COLOR: #f5feed;
font-size:12px;
}
.tx1 { height: 20px;font-size: 9pt; border: 1px solid; border-color: #000000; color: #0000FF}
-->
</style>
</head>
<script language="javascript">
<!--
function ShowProgress() {
	var ProgressID = (new Date()).getTime() % 1000000000;
	var Form = document.MyForm;
	Form.action = "Progress_RunJ.asp?ProgressID=" + ProgressID;
	if (Form.Filedown.value != "") {
		var Ver = navigator.appVersion;
		if (Ver.indexOf("MSIE") > -1 && Ver.substr(Ver.indexOf("MSIE") + 5, 1) > 4) {
			window.showModelessDialog("Progress_WinJ.asp?Count=0&ProgressID=" + ProgressID, null, "dialogWidth=360px; dialogHeight:180px; help:no; status:no");
		} 
		else 
		{
			window.open("Progress_WinJ.asp?Count=0&ProgressID=" + ProgressID, "_blank", "left=240,top=240,width=360,height=160");
		}
		return true;
	} 
	else
	{
	    alert("请选择要上传的文件");
		return false;
	}
}
//-->
</script>
<script>
function ck(obj){if(obj.value.length>0){
var af="flv";
if(eval("with(obj.value)if(!/"+af.split(",").join("|")+"/ig.test(substring(lastIndexOf('.')+1,length)))1;")){alert("只允许上传:\n"+af+"类型文件");obj.createTextRange().execCommand('delete')};
}}
</script>
<body>
<%
if EnableUploadFile="Yes" and (session("adminName")<>"" or session("UserName")<>"") then
%>
<form action="Progress_RunJ.asp" method="post" name="MyForm" onSubmit="return ShowProgress();" enctype="multipart/form-data">
  <input name="Filedown" type="FILE" size="20" class="tx1" onpropertychange="ck(this)">
  <input type="submit" name="Submit" value="上传FLV文件" style="border:1px double rgb(88,88,88);font-size:12px">
</form>
<%
end if
%>
</body>
</html>