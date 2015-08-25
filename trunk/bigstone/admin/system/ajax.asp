<!-- #include file="common.asp" -->
<%
dim func:func = post("cmd","") & "()"
eval(func)

function result()
  if get_session("result","") = "1" then
    call unset_session("result")
	echo 1
  end if
end function

function get_name()
  dim dir:dir = post("dir","loose")
  dim fil:fil = post("fil","loose")
  dim result:result = fil
  dim i,path
  for i = 1 to 1000
	path = server.mappath(dir&result)
	dim fso:set fso = createobject("scripting.filesystemobject")
	if fso.fileexists(path) then
	  result = i & "_" & fil
	else
	  exit for
	end if
	set fso = nothing
  next
  echo result
end function

function get_pic()
  echo get_session("img","")
end function

function get_fil()
  echo get_session("fil","")
end function

function set_fil()
  call set_session("fil",post("fil_url",""),"")
  echo 1
end function

function get_version()
  on error resume next
  dim url:url = "http://www.sin"&"siu.com/njb/version.php"
  dim xml_http:set xml_http = server.createobject("microsoft.xmlhttp")
  xml_http.open "get",url,false
  xml_http.send()
  dim file_data:file_data = xml_http.responsebody
  dim str:str = b2s(file_data)
  if not err and left(str,9) = "njb_send:" then echo str else echo ""
end function

function get_admins()
  dim str
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"admin")
  dim record_count:record_count = db.get_record_count()
  if record_count > 0 then
    dim i
	for i = 0 to record_count - 1
	  str = str&" , "&db.sql_result(i,"adm_user")
	next
	str = mid(str,4,len(str) - 1)
  end if
  echo str
end function

%>