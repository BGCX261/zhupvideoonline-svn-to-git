<%
function get_plu_sheet()
  dim fil,arr()
  dim path:path = server.mappath(S_ROOT&"plugins/")
  dim fso:set fso = server.createobject("scripting.filesystemobject")
  dim folder:set folder = fso.getfolder(path)
  dim files:set files = folder.subfolders
  dim i:i = 0
  for each fil in files
	redim preserve arr(i)
	arr(i) = fil.name
	i = i + 1
  next
  set fso = nothing
  set folder = nothing
  set files = nothing
  if i = 0 then redim preserve arr(-1)
  get_plu_sheet = arr
end function

function get_plu_text(val)
  dim path:path = server.mappath(S_ROOT&"plugins/"&val&"/readme.txt")
  dim fso:set fso = server.createobject("scripting.filesystemobject")
  if fso.fileexists(path) then
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
	set stm = nothing
  else
	txt = "none"
  end if
  set fso = nothing
  get_plu_text = txt
end function

function get_plu_config(val)
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"config where con_name = 'plugin' and con_value = '"&val&"'")
  if db.get_record_count() > 0 then
	get_plu_config = true
  else
	get_plu_config = false
  end if
end function

function get_new_page()
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"config where con_name='new_page'")
  dim record_count:record_count = db.get_record_count()
  dim i,arr()
  if record_count > 0 then
	redim preserve arr(record_count - 1,3)
	for i = 0 to record_count - 1
	  dim val:val = split(db.sql_result(i,"con_value"),"{v}")
	  arr(i,0) = val(0)
	  arr(i,1) = val(1)
	  arr(i,2) = val(2)
	  arr(i,3) = db.sql_result(i,"con_id")
	next
  else
	redim preserve arr(-1)
  end if
  get_new_page = arr
end function

function get_title()
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"config where con_name='title' order by con_id asc")
  dim record_count:record_count = db.get_record_count()
  dim i,arr()
  if record_count > 0 then
	redim preserve arr(record_count - 1,3)
	for i = 0 to record_count - 1
	  dim val:val = split(db.sql_result(i,"con_value"),"{v}")
	  arr(i,0) = val(0)
	  arr(i,1) = val(1)
	  arr(i,2) = val(2)
	  arr(i,3) = db.sql_result(i,"con_id")
	next
  else
	redim preserve arr(-1)
  end if
  get_title = arr
end function

function get_modules_option()
  dim i,arr()
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"modules where mod_action = 'global' order by mod_id asc")
  dim record_count:record_count = db.get_record_count()
  if record_count > 0 then
	redim preserve arr(record_count - 1,1)
	for i = 0 to record_count - 1
	  dim val:val = split(db.sql_result(i,"mod_config"),"|")
	  arr(i,0) = db.sql_result(i,"mod_type")
	  arr(i,1) = val(0)
	next
  else
	redim preserve arr(-1)
  end if
  get_modules_option = arr
end function

function get_modules_sheet()
  dim i,arr()
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"modules order by mod_id asc")
  dim record_count:record_count = db.get_record_count()
  if record_count > 0 then
	redim preserve arr(record_count - 1,4)
	for i = 0 to record_count - 1
	  arr(i,0) = db.sql_result(i,"mod_id")
	  arr(i,1) = db.sql_result(i,"mod_type")
	  arr(i,2) = db.sql_result(i,"mod_action")
	  arr(i,3) = db.sql_result(i,"mod_config")
	  arr(i,4) = db.sql_result(i,"mod_index")
	next
  else
	redim preserve arr(-1)
  end if
  get_modules_sheet = arr
end function

'新秀
%>