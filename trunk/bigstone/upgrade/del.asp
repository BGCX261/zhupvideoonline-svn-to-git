<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>升级至sinsiu 1.2 beta1</title>
</head>
<body>
<%
set fso = createobject("scripting.filesystemobject")
fso.deletefolder server.mappath("new")
fso.deletefile server.mappath("deal.asp")
fso.deletefile server.mappath("index.asp")
set fso = nothing
%>
删除文件成功，但无法完全删除，因为本程序不能删除自身。剩下的文件不会对网站安全构成威胁，若要彻底删除，请手动删除upgrade文件夹
</body>
</html>