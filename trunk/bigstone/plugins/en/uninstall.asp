<!-- #include file="../../admin/system/common.asp" -->
<%
if S_DB_PATH = "plugins/en/" then response.end
on error resume next
dim path:path = server.mappath(S_ROOT&"include/interface.asp")
dim stm:set stm = server.createobject("adodb.stream")
with stm
  .type = 2
  .mode = 3
  .charset = "utf-8"
  .open
  .loadfromfile path
  dim text:text = .readtext
  .close
end with

dim del_text
del_text = chr(60)&"!-- #include file=""../plugins/en/common.asp"" --"&chr(62)&chr(13)&chr(10)
text = replace(text,del_text,"")
del_text = ":if val = ""en"" then call plu_en_reset()"
text = replace(text,del_text,"")
del_text = ":call plu_en_select(val)"
text = replace(text,del_text,"")

dim fso:set fso = server.createobject("scripting.filesystemobject")
fso.deleteFile(path)

set fso = nothing
with stm
  .open
  .writetext text
  .savetofile path,2
  .close
end with
set stm = nothing

call db.reset()
call db.sql_query("delete from config where con_name = 'plugin' and con_value = 'en'")
call db.close_db()
set db = nothing

set fso = server.createobject("scripting.filesystemobject")
fso.deletefile server.mappath("#data_en.mdb")
set fso = nothing

if err then
  err.clear
  echo 0
else
  session("result") = 1
  echo 1
end if

'新秀
%>