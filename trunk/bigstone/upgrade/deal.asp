<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>升级至sinsiu 1.2 beta1</title>
</head>
<body>
<%
database = "../data/#data.mdb"
connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.mappath(database)
set conn = server.createobject("adodb.connection")
conn.open connstr

conn.execute("update config set con_value = 'sinsiu 1.2 beta1' where con_name = 'version'")

set conn = nothing

set fso = createobject("scripting.filesystemobject")
fso.deletefolder server.mappath("../admin")
fso.copyfolder server.mappath("new/admin"),server.mappath("../admin")

fso.copyfolder server.mappath("new/include"),server.mappath("../include")
fso.copyfolder server.mappath("new/templates"),server.mappath("../templates")
set fso = nothing

'修改配置文件
dim db_name,upperbound,lowerbound
Randomize
upperbound = 99999
lowerbound = 10000
db_name = int((upperbound - lowerbound + 1) * rnd + lowerbound)
db_name = db_name & int((upperbound - lowerbound + 1) * rnd + lowerbound)
db_name = db_name & int((upperbound - lowerbound + 1) * rnd + lowerbound)
db_name = db_name & int((upperbound - lowerbound + 1) * rnd + lowerbound)
db_name = "#"&db_name&".mdb"
path = server.mappath("../include/config.asp")
set stm = server.createobject("adodb.stream")
with stm
.type = 2
.mode = 3
.charset = "utf-8"
.open
.loadfromfile path
config = .readtext
.close
end with
config = replace(config,"S_DB_NAME = "&chr(34)&"#data.mdb"&chr(34),"S_DB_NAME = "&chr(34)&db_name&chr(34))
set fso = server.createobject("scripting.filesystemobject")
fso.deletefile(path)
set fso = nothing
with stm
.open
.writetext config
.savetofile path,2
.close
end with
set stm = nothing

set fso = createobject("scripting.filesystemobject")
fso.copyfile server.mappath("../data/#data.mdb"),server.mappath("../data/"&db_name)
fso.deletefile server.mappath("../data/#data.mdb")
set fso = nothing
  
%>
升级成功，为了避免留下安全隐患，<a href="del.asp">请点击这里删除升级文件</a>
</body>
</html>