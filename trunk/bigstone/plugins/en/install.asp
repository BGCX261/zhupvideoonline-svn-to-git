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

dim old_text,add_text,new_text
old_text = chr(60)&chr(37)&"'include"&chr(37)&chr(62)
add_text = chr(60)&"!-- #include file=""../plugins/en/common.asp"" --"&chr(62)&chr(13)&chr(10)
if instr(text,add_text) = 0 then
  new_text = add_text & old_text
  text = replace(text,old_text,new_text)
end if
old_text = "'reset_config"
add_text = ":if val = ""en"" then call plu_en_reset()"
if instr(text,add_text) = 0 then
  new_text = add_text & old_text
  text = replace(text,old_text,new_text)
end if
old_text = "'admin_header"
add_text = ":call plu_en_select(val)"
if instr(text,add_text) = 0 then
  new_text = add_text & old_text
  text = replace(text,old_text,new_text)
end if

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
call db.sql_query("insert into config (con_name,con_value) values ('plugin','en')")
call db.close_db()
set db = nothing

set fso = server.createobject("scripting.filesystemobject")
fso.copyfile server.mappath("../../"&S_DB_PATH&S_DB_NAME),server.mappath("#data_en.mdb")
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