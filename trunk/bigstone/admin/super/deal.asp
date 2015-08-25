<!-- #include file="../system/common.asp" -->
<%
dim func:func = post("cmd","") & "()"
eval(func)

function edit_title()
  dim i,site,title,priority,id,val
  for i = 1 to request.form("site").count
	site = strict(request.form("site")(i))
	title = strict(request.form("title")(i))
	priority = strict(request.form("priority")(i))
	id = strict(request.form("id")(i))
	val = site&"{v}"&title&"{v}"&int(priority)
	call db.reset()
	call db.sql_query("update "&S_DB_PREFIX&"config set con_value='"&val&"' where con_id="&id)
  next
  call set_session("result",1,"")
end function

function add_title()
  dim site:site = post("site","")
  dim title:title = post("title","")
  dim priority:priority = post("priority","")
  dim val:val = site&"{v}"&title&"{v}"&int(priority)
  call db.reset()
  call db.sql_query("insert into "&S_DB_PREFIX&"config (con_name,con_value) values ('title','"&val&"')")
  call set_session("result",1,"")
  echo 1
end function

function del_title()
  dim id:id = post("id","")
  call db.reset()
  call db.sql_query("delete from "&S_DB_PREFIX&"config where con_id="&id)
  call set_session("result",1,"")
  echo 1
end function

function add_module()
  dim val_1:val_1 = post("val_1","")
  dim val_2:val_2 = post("val_2","")
  dim val_3:val_3 = post("val_3","")
  dim val_4:val_4 = post("val_4","")
  dim val_5:val_5 = post("val_5","")
  dim val:val = val_2&"|"&val_3&"|"&val_4&"|"&val_5
  call db.reset()
  call db.sql_query("insert into "&S_DB_PREFIX&"modules (mod_type,mod_action,mod_config) values ('"&val_1&"','global','"&val&"')")
  call set_session("result",1,"")
  echo 1
end function

function add_mod_sheet()
  dim val_1:val_1 = post("val_1","")
  dim val_2:val_2 = post("val_2","")
  dim val_3:val_3 = post("val_3","")
  dim val_4:val_4 = post("val_4","")
  dim val_5:val_5 = post("val_5","")
  dim val:val = val_2&"|"&val_3&"|"&val_4&"|"&val_5
  call db.reset()
  call db.sql_query("insert into "&S_DB_PREFIX&"modules (mod_type,mod_action,mod_config) values ('"&val_1&"_sheet','main','"&val&"')")
  call set_session("result",1,"")
  echo 1
end function

function add_mod_form()
  dim val_1:val_1 = post("val_1","")
  dim val_2:val_2 = post("val_2","")
  dim val_3:val_3 = post("val_3","")
  dim val_4:val_4 = post("val_4","")
  dim val_5:val_5 = post("val_5","")
  dim val_6:val_6 = post("val_6","")
  dim val:val = val_2&"|"&val_3&"|"&val_4&"|"&val_5&"|"&val_6
  call db.reset()
  call db.sql_query("insert into "&S_DB_PREFIX&"modules (mod_type,mod_action,mod_config) values ('"&val_1&"_form','main','"&val&"')")
  call set_session("result",1,"")
  echo 1
end function

function del_module()
  dim id:id = post("id","")
  call db.reset()
  call db.sql_query("delete from "&S_DB_PREFIX&"modules where mod_id="&id)
  call set_session("result",1,"")
  echo 1
end function

function edit_module()
  dim mod_id:mod_id = post("mod_id","")
  dim mod_type:mod_type = post("mod_type","")
  dim mod_action:mod_action = post("mod_action","")
  dim mod_config:mod_config = post("mod_config","")
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"modules set mod_type='"&mod_type&"',mod_action='"&mod_action&"',mod_config='"&mod_config&"' where mod_id="&mod_id)
  call set_session("result",1,"")
  echo 1
end function

function set_mod_index()
  dim mod_id:mod_id = post("id","")
  dim mod_index:mod_index = post("index","")
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"modules set mod_index = '"&mod_index&"' where mod_id = "&mod_id)
  echo 1
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

'新秀
%>