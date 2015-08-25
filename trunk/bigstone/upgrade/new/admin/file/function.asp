<%
function get_tpl_sheet()
  dim fil,arr()
  dim path:path = server.mappath(S_ROOT&"templates/")
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
  get_tpl_sheet = arr
end function

function get_tpl_text(val)
  dim path:path = server.mappath(S_ROOT&"templates/"&val&"/readme.txt")
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
  get_tpl_text = txt
end function

function get_tpl_files()
  dim fil,arr()
  dim path:path = server.mappath(S_TPL_PATH)
  dim fso:set fso = server.createobject("scripting.filesystemobject")
  dim folder:set folder = fso.getfolder(path)
  dim files:set files = folder.files
  dim i:i = 0
  for each fil in files
	redim preserve arr(i)
	arr(i) = fil.name
	i = i + 1
  next
  path = server.mappath(S_TPL_PATH&"module/")
  set fso = server.createobject("scripting.filesystemobject")
  set folder = fso.getfolder(path)
  set files = folder.files
  for each fil in files
	redim preserve arr(i)
	arr(i) = "module/"&fil.name
	i = i + 1
  next
  path = server.mappath(S_TPL_PATH&"js/")
  set fso = server.createobject("scripting.filesystemobject")
  set folder = fso.getfolder(path)
  set files = folder.files
  for each fil in files
	redim preserve arr(i)
	arr(i) = "js/"&fil.name
	i = i + 1
  next
  set fso = nothing
  set folder = nothing
  set files = nothing
  if i = 0 then redim preserve arr(-1)
  get_tpl_files = arr
end function

function get_folder_sheet(val)
  dim fol,arr()
  dim path:path = server.mappath(val)
  dim fso:set fso = server.createobject("scripting.filesystemobject")
  dim folder:set folder = fso.getfolder(path)
  dim folders:set folders = folder.Subfolders
  dim i:i = 0
  for each fol in folders
	redim preserve arr(i)
	arr(i) = fol.name
	i = i + 1
  next
  set fso = nothing
  set folder = nothing
  set folders = nothing
  if i = 0 then redim preserve arr(-1)
  get_folder_sheet = arr
end function

function get_file_sheet(val)
  dim fil,arr()
  dim path:path = server.mappath(val)
  dim fso:set fso = server.createobject("scripting.filesystemobject")
  dim folder:set folder = fso.getfolder(path)
  dim files:set files = folder.files
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
  get_file_sheet = arr
end function

function sel_file(val)
  dim return
  return = get_id("article","art_img",val)
  if return = false then return = get_id("article","art_reduce_img",val)
  if return = false then return = get_id("picture","pic_path",val)
  sel_file = return
end function

function get_focus_sheet()
  dim i,arr()
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"picture where pic_type = 'focus' order by pic_top desc,pic_index desc,pic_id asc")
  dim record_count:record_count = db.get_record_count()
  if record_count > 0 then
	redim preserve arr(record_count - 1,7)
	for i = 0 to record_count - 1
	  arr(i,0) = db.sql_result(i,"pic_id")
	  arr(i,1) = db.sql_result(i,"pic_path")
	  arr(i,2) = db.sql_result(i,"pic_url")
	  arr(i,3) = db.sql_result(i,"pic_title")
	  arr(i,4) = db.sql_result(i,"pic_site")
	  arr(i,5) = db.sql_result(i,"pic_index")
	  arr(i,6) = db.sql_result(i,"pic_top")
	  arr(i,7) = db.sql_result(i,"pic_show")
	next
  else
	redim preserve arr(-1)
  end if
  get_focus_sheet = arr
end function

function get_lang_sheet()
  dim fil,arr()
  dim path:path = server.mappath(S_ROOT&"languages/")
  dim fso:set fso = server.createobject("scripting.filesystemobject")
  dim folder:set folder = fso.getfolder(path)
  dim files:set files = folder.files
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
  get_lang_sheet = arr
end function

'新秀
%>