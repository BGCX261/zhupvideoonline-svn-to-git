<!-- #include file="../system/common.asp" -->
<%
dim func:func = post("cmd","") & "()"
eval(func)

function sel_tpl()
  dim tpl:tpl = post("tpl","")
  dim n:n = "S_TPL_PATH"
  dim c:c = "S_ROOT&"&chr(34)&right(S_TPL_PATH,len(S_TPL_PATH) - len(S_ROOT))&chr(34)
  dim v:v = "S_ROOT&"&chr(34)&"templates/"&tpl&"/"&chr(34)
  call edit_config(n,c,v)
  call set_session("result",1,"")
  echo 1
end function

function edit_tpl()
  dim fil:fil = post("fil","")
  dim tpl:tpl = request.form("tpl")
  dim path:path = server.mappath(S_TPL_PATH&fil)
  dim stm:set stm = server.createobject("adodb.stream")
  with stm
	.type = 2
	.mode = 3
	.charset = "utf-8"
	.open
	.writetext tpl
	.savetofile path,2
	.close
  end with
  set stm = nothing
  call set_session("result",1,"")
end function

function del_pic()
  dim fil:fil = post("id","")
  dim path:path = server.mappath(fil)
  dim fso:set fso = server.createobject("scripting.filesystemobject")
  if fso.fileexists(path) then
	fso.deleteFile(path)
  end if
  set fso = nothing
  call set_session("result",1,"")
  echo 1
end function

function edit_focus()
  dim foc_id:foc_id = post("foc_id","")
  dim pic_path:pic_path = post("pic_path","")
  dim pic_url:pic_url = post("pic_url","")
  dim pic_title:pic_title = post("pic_title","")
  dim pic_site:pic_site = post("pic_site","")
  if pic_url = "" or pic_url = "http://" then pic_url = "index.asp"
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"picture set pic_path = '"&pic_path&"',pic_url = '"&pic_url&"',pic_title = '"&pic_title&"',pic_site = '"&pic_site&"' where pic_id="&foc_id)
  call set_session("result",1,"")
  echo 1
end function

function add_focus()
  dim pic_path:pic_path = post("pic_path","")
  dim pic_url:pic_url = post("pic_url","")
  dim pic_title:pic_title = post("pic_title","")
  dim pic_site:pic_site = post("pic_site","")
  if pic_url = "" or pic_url = "http://" then pic_url = "index.asp"
  call db.reset()
  call db.sql_query("insert into "&S_DB_PREFIX&"picture (pic_path,pic_url,pic_title,pic_site,pic_type) values ('"&pic_path&"','"&pic_url&"','"&pic_title&"','"&pic_site&"','focus')")
  call set_session("result",1,"")
  echo 1
end function

function del_focus()
  dim id:id = post("id","")
  call db.reset()
  call db.sql_query("delete from "&S_DB_PREFIX&"picture where pic_id = "&id)
  call set_session("result",1,"")
  echo 1
end function

function sel_lang()
  dim lang:lang = post("lang","")
  dim n:n = "S_LANG"
  dim c:c = chr(34)&S_LANG&chr(34)
  dim v:v = chr(34)&replace(lang,".asp","")&chr(34)
  call edit_config(n,c,v)
  call set_session("result",1,"")
  echo 1
end function

function edit_lang()
  dim fil:fil = post("fil","")
  dim lang:lang = request.form("lang")
  dim path:path = server.mappath(S_ROOT&"languages/"&fil)
  dim stm:set stm = server.createobject("adodb.stream")
  with stm
	.type = 2
	.mode = 3
	.charset = "utf-8"
	.open
	.writetext lang
	.savetofile path,2
	.close
  end with
  set stm = nothing
  call set_session("result",1,"")
end function

function del_file()
  dim fil:fil = post("id","")
  dim path:path = server.mappath(fil)
  dim fso:set fso = server.createobject("scripting.filesystemobject")
  if fso.fileexists(path) then
	fso.deleteFile(path)
  end if
  set fso = nothing
  call set_session("result",1,"")
  echo 1
end function

function set_foc_index()
  dim pic_id:pic_id = post("id","")
  dim pic_index:pic_index = post("index","")
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"picture set pic_index = '"&pic_index&"' where pic_id = "&pic_id)
  echo 1
end function

function set_foc_top()
  dim pic_id:pic_id = post("id","")
  dim pic_top:pic_top = post("top","")
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"picture set pic_top = "&pic_top&" where pic_id = "&pic_id)
  call set_session("result",1,"")
  echo 1
end function

function set_foc_show()
  dim pic_id:pic_id = post("id","")
  dim pic_show:pic_show = post("show","")
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"picture set pic_show = "&pic_show&" where pic_id = "&pic_id)
  call db.reset()
  echo 1
end function

call db.close_db()
set db = nothing
'新秀
%>