<!-- #include file="../system/common.asp" -->
<%
dim func:func = post("cmd","") & "()"
eval(func)

function edit_contact()
  dim i,cue,content,id,val
  for i = 1 to request.form("cue").count
	cue = strict(request.form("cue")(i))
	content = strict(request.form("content")(i))
	id = strict(request.form("id")(i))
	val = cue&"{v}"&content
	call db.reset()
	call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&val&"' where con_id="&id)
  next
  call set_session("result",1,"")
end function

function add_contact()
  dim cue:cue = post("cue","")
  dim content:content = post("content","")
  dim val:val = cue&"{v}"&content
  call db.reset()
  call db.sql_query("insert into "&S_DB_PREFIX&"config (con_name,con_value) values ('contact','"&val&"')")
  call set_session("result",1,"")
  echo 1
end function

function del_contact()
  dim id:id = post("id","")
  call db.reset()
  call db.sql_query("delete from "&S_DB_PREFIX&"config where con_id="&id)
  call set_session("result",1,"")
  echo 1
end function

function edit_site()
  dim sit_title:sit_title = post("sit_title","")
  dim sit_name:sit_name = post("sit_name","")
  dim sit_domain:sit_domain = post("sit_domain","")
  dim sit_code:sit_code = post("sit_code","")
  dim sit_code_url:sit_code_url = post("sit_code_url","")
  dim sit_tec:sit_tec = post("sit_tec","")
  dim sit_tec_url:sit_tec_url = post("sit_tec_url","")
  dim sit_keywords:sit_keywords = post("sit_keywords","")
  dim sit_description:sit_description = post("sit_description","")
  dim sit_count_code:sit_count_code = post("sit_count_code","loose")
  sit_description = left(sit_description,250)
  sit_count_code = left(sit_count_code,250)
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&sit_title&"' where con_name='sit_title'")
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&sit_name&"' where con_name='sit_name'")
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&sit_domain&"' where con_name='sit_domain'")
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&sit_code&"' where con_name='sit_code'")
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&sit_code_url&"' where con_name='sit_code_url'")
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&sit_tec&"' where con_name='sit_tec'")
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&sit_tec_url&"' where con_name='sit_tec_url'")
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&sit_keywords&"' where con_name='sit_keywords'")
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&sit_description&"' where con_name='sit_description'")
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&sit_count_code&"' where con_name='sit_count_code'")
  call set_session("result",1,"")
end function

function edit_nav()
  dim i,word,url,id,val
  for i = 1 to request.form("word").count
	word = strict(request.form("word")(i))
	url = strict(request.form("url")(i))
	id = strict(request.form("id")(i))
	call db.reset()
	call db.sql_query("update "&S_DB_PREFIX&"menu set men_name='"&word&"',men_url='"&url&"' where men_id="&id)
  next
  call set_session("result",1,"")
end function

function add_nav()
  dim men_type:men_type = post("men_type","")
  dim word:word = post("word","")
  dim url:url = post("url","")
  url = replace(url,"{njb}","&")
  call db.reset()
  call db.sql_query("insert into "&S_DB_PREFIX&"menu (men_type,men_name,men_url) values ('"&men_type&"','"&word&"','"&url&"')")
  call set_session("result",1,"")
  echo 1
end function

function del_nav()
  dim id:id = post("id","")
  call db.reset()
  call db.sql_query("delete from "&S_DB_PREFIX&"menu where men_id="&id)
  call set_session("result",1,"")
  echo 1
end function

function set_nav_index()
  dim men_id:men_id = post("id","")
  dim men_index:men_index = post("index","")
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"menu set men_index = '"&men_index&"' where men_id = "&men_id)
  echo 1
end function

function set_nav_top()
  dim men_id:men_id = post("id","")
  dim men_top:men_top = post("top","")
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"menu set men_top = "&men_top&" where men_id = "&men_id)
  call set_session("result",1,"")
  echo 1
end function

function set_nav_show()
  dim men_id:men_id = post("id","")
  dim men_show:men_show = post("show","")
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"menu set men_show = "&men_show&" where men_id = "&men_id)
  call db.reset()
  echo 1
end function

function edit_safe()
  dim old_user_name:old_user_name = post("old_user_name","")
  dim old_password:old_password = md5(post("old_password",""))
  dim new_user_name:new_user_name = post("new_user_name","")
  dim new_password:new_password = md5(post("new_password",""))
  dim id
  if old_user_name <> "" and new_user_name <> "" then
	call db.reset()
	call db.sql_query("select * from "&S_DB_PREFIX&"admin where adm_user='"&old_user_name&"' and adm_password='"&old_password&"'")
	if db.get_record_count() > 0 then
	  id = db.sql_result(0,"adm_id")
	  call db.sql_query("update "&S_DB_PREFIX&"admin set adm_user='"&new_user_name&"',adm_password='"&new_password&"' where adm_id="&id)
	  if old_user_name = get_session("user","") then
		unset_session("login")
		unset_session("user")
		unset_session("password")
		echo 1
	  else
		call set_session("result",1,"")
		echo 2
	  end if
	else
	  echo 3
	end if
  elseif old_user_name = "" and new_user_name <> "" then
	call db.reset()
	call db.sql_query("select * from "&S_DB_PREFIX&"admin where adm_user='"&new_user_name&"'")
	if db.get_record_count() = 0 then
	  call db.sql_query("insert into "&S_DB_PREFIX&"admin (adm_user,adm_password) values ('"&new_user_name&"','"&new_password&"')")
	  call set_session("result",1,"")
	  echo 4
	else
	  echo 5
	end if
  elseif old_user_name <> "" and new_user_name = "" then
	if old_user_name <> get_session("user","") then
	  call db.reset()
	  call db.sql_query("select * from "&S_DB_PREFIX&"admin where adm_user='"&old_user_name&"' and adm_password='"&old_password&"'")
	  if db.get_record_count() > 0 then
		id = db.sql_result(0,"adm_id")
		call db.sql_query("delete from "&S_DB_PREFIX&"admin where adm_id="&id)
		call set_session("result",1,"")
		echo 6
	  else
		echo 7
	  end if
	else
	  echo 8
	end if
  end if
end function

function edit_show()
  dim id:id = post("id","")
  dim arr:arr = split(id,"-")
  dim sinty_xml:sinty_xml = server.mappath(S_TPL_PATH&"sinty.xml")
  dim xml:set xml = server.createobject("microsoft.xmldom")
  xml.load(sinty_xml)
  dim len_:len_ = xml.selectsinglenode("site").selectsinglenode(arr(0)).childnodes.length
  dim i,text
  for i = 0 to len_ - 1
	text = xml.selectsinglenode("site").selectsinglenode(arr(0)).childnodes.item(i).text
	if arr(1) = replace(text,"*","") then
	  if arr(2) = 1 then
		xml.selectsinglenode("site").selectsinglenode(arr(0)).childnodes.item(i).text = replace(text,"*","")
		call edit_tpl_show(arr(0),replace(text,"*",""),1)
	  else
		xml.selectsinglenode("site").selectsinglenode(arr(0)).childnodes.item(i).text = "*" & replace(text,"*","")
		call edit_tpl_show(arr(0),replace(text,"*",""),0)
	  end if
	  xml.save(sinty_xml)
	end if
  next
  set xml = nothing
  echo 1
end function

function edit_tpl_show(fil,module,show)
  dim path:path = server.mappath(S_TPL_PATH&fil&".asp")
  dim stm:set stm = server.createobject("adodb.stream")
  with stm
	.type = 2
	.mode = 3
	.charset = "utf-8"
	.open
	.loadfromfile path
	dim txt:txt = .readtext
	.close
  end with
  if show = 0 then
	txt = replace(txt,"{inc:module/"&module&"}","{*inc:module/"&module&"*}")
  else
	txt = replace(txt,"{*inc:module/"&module&"*}","{inc:module/"&module&"}")
  end if
  dim fso:set fso = server.createobject("scripting.filesystemobject")
  fso.deleteFile(path)
  set fso = nothing
  with stm
	.open
	.writetext txt
	.savetofile path,2
	.close
  end with
  set stm = nothing
end function

function backup_db()
  dim fil:fil = "#"&get_time()&".asp"
  dim fso:set fso = createobject("scripting.filesystemobject")
  dim path_old:path_old = server.mappath("../../data")&"/"&S_DB_NAME
  dim path_new:path_new = server.mappath("../../data/backup")&"/"&fil
  fso.copyfile path_old,path_new
  set fso = nothing
  call set_session("result",1,"")
  echo 1
end function

function restore_db()
  dim fil:fil = post("id","")
  if right(fil,4) = ".asp" then
	dim path_new:path_new = server.mappath("../../data/backup")&"/"&fil
	dim path_old:path_old = server.mappath("../../data")&"/"&S_DB_NAME
	dim fso:set fso = createobject("Scripting.filesystemobject")
	if fso.fileexists(path_new) then
	  set db = nothing
	  fso.deleteFile(path_old)
	  fso.copyfile path_new,path_old
	end if
	set fso = nothing
	call set_session("result",1,"")
	echo 1
  end if
end function

function del_db()
  dim fil:fil = post("id","")
  dim db_path:db_path = server.mappath("../../data/backup")&"/"&fil
  dim fso:set fso = createobject("Scripting.filesystemobject")
  if fso.fileexists(db_path) then fso.deletefile(db_path)
  set fso = nothing
  call set_session("result",1,"")
  echo 1
end function

function set_about_filter()
  dim val:val = int(post("val",""))
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&val&"' where con_name='about_filter'")
  echo 1
end function

function set_about_len()
  dim val:val = int(post("val",""))
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&val&"' where con_name='about_len'")
  echo 1
end function

function set_list_len()
  dim con_name:con_name = post("con_name","")
  dim val:val = int(post("val",""))
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&val&"' where con_name='"&con_name&"'")
  echo 1
end function

function set_reduce_img()
  dim con_name:con_name = post("con_name","")
  dim val:val = int(post("val",""))
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&val&"' where con_name='"&con_name&"'")
  echo 1
end function

function set_session_method()
  dim val:val = int(post("val",""))
  call edit_config("S_SESSION",S_SESSION,val)
  echo 1
end function

function set_compile()
  dim val:val = int(post("val",""))
  dim sinty_xml:sinty_xml = server.mappath(S_TPL_PATH&"sinty.xml")
  dim xml:set xml = server.createobject("microsoft.xmldom")
  xml.load(sinty_xml)
  xml.selectsinglenode("site").setattribute "compile",val
  xml.save(sinty_xml)
  set xml = nothing
  echo 1
end function

function set_static()
  dim val:val = int(post("val",""))
  call edit_config("S_STATIC",S_STATIC,val)
  echo 1
end function

function clear_static()
  dim prefix:prefix = post("prefix","")
  dim path:path = server.mappath(S_ROOT&"html")
  dim fso:set fso = createobject("scripting.filesystemobject")
  if prefix = "0" then
	fso.deletefolder path
	fso.createfolder path
  else
    dim fil
	dim folder:set folder = fso.getfolder(path)
	dim files:set files = folder.files
	for each fil in files
	  if(left(fil.name,len(prefix)) = prefix) then
		fso.deletefile server.mappath(S_ROOT&"html/"&fil.name)
	  end if
	next
  end if
  echo 1
end function

function set_sentmail()
  dim val:val = int(post("val",""))
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&val&"' where con_name='sentmail'")
  echo 1
end function

function save_sentmail()
  dim smtp:smtp = post("smtp","")
  dim send:send = post("send","")
  dim password:password = post("password","")
  dim receive:receive = post("receive","")
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&smtp&"' where con_name='sentmail_smtp'")
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&send&"' where con_name='sentmail_send'")
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&password&"' where con_name='sentmail_password'")
  call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&receive&"' where con_name='sentmail_receive'")
  echo 1
end function

if post("cmd","") <> "restore_db" then
  call db.close_db()
  set db = nothing
end if

'新秀
%>