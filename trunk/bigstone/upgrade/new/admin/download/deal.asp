<!-- #include file="../system/common.asp" -->
<%
dim sql_1,sql_2,sql_3,att,title
dim func:func = post("cmd","") & "()"
eval(func)

function add_art()
  sql_1 = ""
  sql_2 = ""
  att = ""
  dim module:module = post("module","")
  dim rank:rank = get_module_config(module&"_form")
  for i = 0 to ubound(rank)
	call set_deal_add(rank(i,1))
  next
  call set_deal_add("art_keywords|art_keywords||befort_form_text|deal_form_text")
  call set_deal_add("art_description|art_description||befort_form_text|deal_form_text")
  if len(att) > 3 then att = left(left(att,len(att)-3),250)
  dim art_add_time:art_add_time = now()
  call db.reset()
  call db.sql_query("insert into "&S_DB_PREFIX&"article (art_type,"&sql_1&"art_attributes,art_add_time) values ('"&module&"',"&sql_2&"'"&att&"','"&art_add_time&"')")
  call set_art_nav(module)
  call unset_session("img")
  call set_session("result",1,"")
  echo 1
end function

function set_deal_add(rank)
  dim arr:arr = split(rank,"|")
  dim func:func = arr(4) & "("&chr(34)&arr(0)&chr(34)&","&chr(34)&arr(1)&chr(34)&")"
  eval(func)
end function 

function deal_form_text(a,b)
  dim val:val = post(b,"")
  if a = "art_title" then title = val
  sql_1 = sql_1 & a & ","
  sql_2 = sql_2 & "'" & val & "',"
  sql_3 = sql_3 & a & "='" & val & "',"
end function

function deal_form_img(a,b)
  dim art_img:art_img = post("pic_path","")
  dim art_reduce_img
  if art_img <> "" and instr(art_img,"/no_img.gif") = 0 then
	art_img = mid(art_img,len(S_ROOT) + 1,len(art_img) - len(S_ROOT))
	dim img_name:img_name = get_file_name(art_img,"/")
	art_reduce_img = replace(art_img,img_name,"reduce_"&img_name)
  else
	art_img = "images/no_img.gif"
	art_reduce_img = "images/no_img.gif"
  end if
  sql_1 = sql_1 & "art_img" & ","
  sql_2 = sql_2 & "'" & art_img & "',"
  sql_3 = sql_3 & "art_img='" & art_img & "',"
  sql_1 = sql_1 & "art_reduce_img" & ","
  sql_2 = sql_2 & "'" & art_reduce_img & "',"
  sql_3 = sql_3 & "art_reduce_img='" & art_reduce_img & "',"
end function

function deal_form_file(a,b)
  dim fil_path:fil_path = post("fil_path","")
  if fil_path <> "" and left(fil_path,1) = "/" then fil_path = mid(fil_path,len(S_ROOT) + 1,len(fil_path) - len(S_ROOT))
  att = att & "att_url:" & replace(fil_path,"http://","http;//") & "{v}"
end function

function deal_form_editor(a,b)
  dim val:val = post(b,"loose")
  if a = "art_text" then
    dim imgs:imgs = ""
	dim text:text = ""
	dim module:module = post("module","")
	dim mod_config:mod_config = get_module_config(module)
	dim arr:arr = split(mod_config(0,1),"|")
	if ubound(arr) >= 0 then
	  if arr(2) <> "0" then imgs = get_img_in_text(val)
	  if arr(3) <> "0" then text = cut_str(remove_html(val),arr(3))
	end if
	sql_1 = sql_1 & a & ",art_more_img,art_short_text,"
	sql_2 = sql_2 & "'" & val & "','" & imgs & "','" & text & "',"
	sql_3 = sql_3 & a & "='" & val & "',art_more_img='" & imgs & "',art_short_text='" & text & "',"
  else
	sql_1 = sql_1 & a & ","
	sql_2 = sql_2 & "'" & val & "',"
	sql_3 = sql_3 & a & "='" & val & "',"
  end if
end function

function deal_form_att(a,b)
  dim val:val = post(b,"")
  att = att & a & ":" & val & "{v}"
end function

function set_art_nav(module)
  on error resume next
  dim art_id:art_id = get_new_data("article","art_id")
  dim mod_config:mod_config = get_module_config(module)
  dim arr:arr = split(mod_config(0,1),"|")
  if title <> "" and art_id <> false and arr(1) <> "none" then
    call db.reset()
	call db.sql_query("insert into "&S_DB_PREFIX&"menu (men_type,men_name,men_url) values ('"&arr(1)&"','"&title&"','?"&module&"_"&art_id&"."&S_SUFFIX&"')")
  end if
  if err then err.clear
end function

function edit_art()
  sql_1 = ""
  sql_2 = ""
  att = ""
  dim module:module = post("module","")
  dim id:id = post("id","")
  dim rank:rank = get_module_config(module&"_form")
  for i = 0 to ubound(rank)
	call set_deal_add(rank(i,1))
  next
  call set_deal_add("art_keywords|art_keywords||befort_form_text|deal_form_text")
  call set_deal_add("art_description|art_description||befort_form_text|deal_form_text")
  if len(att) > 3 then att = left(left(att,len(att)-3),250)
  dim art_add_time:art_add_time = now()
  dim sql:sql = "update "&S_DB_PREFIX&"article set "&sql_3&"art_add_time='"&art_add_time&"',art_attributes='"&att&"' where art_id="&id
  call db.reset()
  call db.sql_query(sql)
  call unset_session("img")
  call set_session("result",1,"")
  echo 1
end function

function del_art()
  dim arr:arr = split(post("id",""),",")
  dim id:id = arr(0)
  dim page
  if arr(1) = "" then page = 1 else page = arr(1)
  call db.reset()
  call db.sql_query("select art_reduce_img,art_img from "&S_DB_PREFIX&"article where art_id = "&id)
  if db.get_record_count() > 0 then
	dim img_reduce_path:img_reduce_path = db.sql_result(0,"art_reduce_img")
	dim img_path:img_path = db.sql_result(0,"art_img")
	call db.sql_query("delete from "&S_DB_PREFIX&"article where art_id="&id)
	dim path,fso
	set fso = createobject("Scripting.filesystemobject")
	path = server.mappath(S_ROOT&img_reduce_path)
	if fso.fileexists(path) and img_reduce_path <> "images/no_img.gif" then fso.deleteFile(path)
	path = server.mappath(S_ROOT&img_path)
	if fso.fileexists(path) and img_path <> "images/no_img.gif" then fso.deleteFile(path)
	set fso = nothing
  end if
  call set_session("result",1,"")
  echo 1
end function

function add_cat()
  dim module:module = post("module","")
  dim cat_parent_id:cat_parent_id = post("cat_parent_id","")
  dim cat_name:cat_name = post("cat_name","")
  call db.reset()
  call db.sql_query("insert into "&S_DB_PREFIX&"category (cat_type,cat_parent_id,cat_name) values ('"&module&"','"&cat_parent_id&"','"&cat_name&"')")
  call set_session("result",1,"")
  echo 1
end function

function edit_cat()
  dim i,cat_parent_id,cat_name,cat_id
  for i = 1 to request.form("cat_id").count
	cat_parent_id = strict(request.form("cat_parent_id")(i))
	cat_name = strict(request.form("cat_name")(i))
	cat_id = strict(request.form("cat_id")(i))
	call db.reset()
	call db.sql_query("update "&S_DB_PREFIX&"category set cat_parent_id="&cat_parent_id&",cat_name='"&cat_name&"' where cat_id="&cat_id)
  next
  call set_session("result",1,"")
end function

function del_cat()
  dim cat_id:cat_id = post("id","")
  call db.reset()
  call db.sql_query("delete from "&S_DB_PREFIX&"category where cat_id="&cat_id)
  call set_session("result",1,"")
  echo 1
end function

function set_art_index()
  dim art_id:art_id = post("id","")
  dim art_index:art_index = post("index","")
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"article set art_index = '"&art_index&"' where art_id = "&art_id)
  echo 1
end function

function set_art_top()
  dim art_id:art_id = post("id","")
  dim art_top:art_top = post("top","")
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"article set art_top = "&art_top&" where art_id = "&art_id)
  echo 1
end function

function set_art_show()
  dim art_id:art_id = post("id","")
  dim art_show:art_show = post("show","")
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"article set art_show = "&art_show&" where art_id = "&art_id)
  call db.reset()
  echo 1
end function

function set_cat_index()
  dim cat_id:cat_id = post("id","")
  dim cat_index:cat_index = post("index","")
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"category set cat_index = '"&cat_index&"' where cat_id = "&cat_id)
  echo 1
end function

function set_cat_top()
  dim cat_id:cat_id = post("id","")
  dim cat_top:cat_top = post("top","")
  call db.reset()
  call db.sql_query("update "&S_DB_PREFIX&"category set cat_top = "&cat_top&" where cat_id = "&cat_id)
  echo 1
end function

function set_cat_show()
  dim i,flag:flag = 1
  dim cat_id:cat_id = post("id","")
  dim cat_show:cat_show = post("show","")
  dim family:family = get_cat_family("category",cat_id)
  dim cat_parent_id:cat_parent_id = get_data("category",int(cat_id),"cat_parent_id")
  if cat_parent_id <> 0 then
	if get_data("category",cat_parent_id,"cat_show") = 0 then flag = 0
  end if
  if flag <> 0 then
	if ubound(family) = 0 then
	  flag = 1
	elseif ubound(family) > 0 then
	  for i = 1 to ubound(family)
		if get_data("category",family(i),"cat_show") = 1 then
		  flag = 2
		  exit for
		end if
	  next
	end if
  end if
  if flag = 1 then
	call db.reset()
	call db.sql_query("update "&S_DB_PREFIX&"category set cat_show = "&cat_show&" where cat_id = "&cat_id)
	call db.reset()
	echo 1
  elseif flag = 2 then
    echo 2
  end if
end function

call db.close_db()
set db = nothing
'新秀
%>