<%
'输出函数
function echo(val)
  response.write(val)
end function
'获取配置值
function get_config(con_name)
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"config where con_name='"&con_name&"'")
  if db.get_record_count() > 0 then
	get_config = db.sql_result(0,"con_value")
  else
	get_config = false
  end if
end function
'获取指定ID的数据
function get_data(table,id,field_)
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&table&" where "&left(table,3)&"_id = "&id)
  if db.get_record_count() > 0 then
	get_data = db.sql_result(0,field_)
  else
	get_data = false
  end if
end function
'获取指定数据的ID
function get_id(table,field_,value_)
  call db.reset()
  if typename(value_) = "String" then
	call db.sql_query("select * from "&S_DB_PREFIX&table&" where "&field_&" = '"&value_&"'")
  else
	call db.sql_query("select * from "&S_DB_PREFIX&table&" where "&field_&" = "&value_)
  end if
  if db.get_record_count() > 0 then
	get_id = db.sql_result(0,left(table,3)&"_id")
  else
	get_id = false
  end if
end function
'获取最新记录指定字段的数据
function get_new_data(table,field_)
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&table&" order by "&left(table,3)&"_id desc")
  if db.get_record_count() > 0 then
	get_new_data = db.sql_result(0,field_)
  else
	get_new_data = false
  end if
end function
'获取指定ID的分类及其所有子类
function get_cat_family(cat_type,id)
  dim i,j,arr(),family(),falg
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"category where cat_type='"&cat_type&"'")
  dim record_count:record_count = db.get_record_count()
  if record_count > 0 then
	redim preserve arr(record_count - 1,1)
	for i = 0 to record_count - 1
	  arr(i,0) = db.sql_result(i,"cat_id")
	  arr(i,1) = db.sql_result(i,"cat_parent_id")
	next
  else
	redim preserve arr(-1)
  end if
  redim preserve family(0)
  family(0) = id & ""
  flag = 1
  do while flag = 1
    flag = 0
	for i = 0 to ubound(family)
	  for j = 0 to ubound(arr)
		if family(i) = arr(j,1) then
		  redim preserve family(ubound(family) + 1)
		  family(ubound(family)) = arr(j,0)
		  arr(j,1) = -1
		  flag = 1
		end if
	  next
	next
  loop
  get_cat_family = family
end function
'获取时间数值串
function get_time()
  dim str
  str = replace(now()," ","")
  str = replace(str,"-","")
  str = replace(str,":","")
  get_time = str
end function
'获取日期数值串
function get_date()
  dim str
  str = replace(date(),"-","")
  str = replace(str,"/","")
  get_date = str
end function
'严格过滤字符串中的危险符号
function strict(byref str)
  str = replace(str,"<","&#60;")
  str = replace(str,">","&#62;")
  str = replace(str,"%","&#37;")
  str = replace(str,chr(39),"&#39;")
  str = replace(str,chr(34),"&#34;")
  str = replace(str,chr(13)&chr(10),"")
  strict = str
end function
'宽松过滤字符串中的危险符号
function loose(byref str)
  str = replace(str,chr(39),"&#39;")
  str = replace(str,chr(60)&chr(37),"")
  str = replace(str,chr(37)&chr(62),"")
  str = replace(str,chr(13)&chr(10),"")
  loose = str
end function
'截取字符串
function cut_str(str,num)
  num = int(num)
  if len(str) > num then
	cut_str = left(str,num)&"..."
  else
	cut_str = str
  end if
end function
'过滤HTML标签
function remove_html(byval html)
  dim reg:set reg = new regexp
  reg.ignorecase = true
  reg.global = true
  reg.pattern = "<.*?>"
  dim mat
  dim matches:set matches = reg.execute(html)
  for each mat in matches
	html = replace(html,mat.value,"")
  next
  reg.pattern = "<.*"
  set matches = reg.execute(html)
  for each mat in matches
	html = replace(html,mat.value,"")
  next
  set reg = nothing
  set matches = nothing
  remove_html = html
end function
'修复HTML标签
function repair_html(byval html)
  html = html & "#end#"
  dim reg:set reg = new regexp
  reg.ignorecase = true
  reg.global = true
  reg.pattern = "<[^>]*?(#end#)"
  set matches = reg.execute(html)
  for each mat in matches
	html = replace(html,mat.value,"")
  next
  html = replace(html,"#end#","")
  set reg = nothing
  set matches = nothing
  repair_html = html
end function
'数字限界
function num_bound(min,max,num)
  if num = "" then num = 0 else num = int(num)
  if min < max then
	if num < min then
	  num = min
	elseif num > max then
	  num = max
	end if
  else
	num = min
  end if
  num_bound = num
end function
'获取完整路径中的文件名或扩展名
function get_file_name(full_path,str)
  if full_path <> "" then
	get_file_name = mid(full_path,instrrev(full_path,str) + 1)
  else
	get_file_name = ""
  end if
end function
'获取本页面URL
function get_url()
  dim url:url = request.servervariables("query_string")
  if left(url,10) = "index.asp?" then url = right(url,len(url) - 10)
  get_url = url
end function
'获取属性列表中的某个属性值
function get_attribute(att_str,att_name)
  dim result:result = ""
  dim arr:arr = split(att_str,"{v}")
  dim i:for i = 0 to ubound(arr)
	dim arr2:arr2 = split(arr(i),":")
	if ubound(arr2) >= 0 then
	  if arr2(0) = att_name then result = arr2(1)
	end if
  next
  get_attribute = result
end function
'设置session
function set_session(nam,val,fil)
  if fil = "" then fil = "strict"
  if S_SESSION = 1 then
	session(nam) = eval(fil&"(val)")
  else
    response.cookies(nam) = eval(fil&"(val)")
  end if
end function
'获取session
function get_session(nam,fil)
  if fil = "" then fil = "strict"
  if S_SESSION = 1 then
	get_session = eval(fil&"(session(nam))")
  else
    get_session = eval(fil&"(request.cookies(nam))")
  end if
end function
'销毁session
function unset_session(nam)
  if S_SESSION = 1 then
	session.contents.remove(nam)
  else
	response.cookies(nam).expires = (now()-1)
  end if
end function
'获取post
function post(val,fil)
  if fil = "" then fil = "strict"
  post = eval(fil&"(request.form(val))")
end function
'获取get
function gets(val,fil)
  if fil = "" then fil = "strict"
  gets = eval(fil&"(request.querystring(val))")
end function
'获取客户端IP
function get_ip()
  dim str
  if request.servervariables("http_x_forwarded_for") = "" or instr(request.servervariables("http_x_forwarded_for"),"unknown") > 0 then
	str = request.servervariables("remote_addr")
  elseif instr(request.servervariables("http_x_forwarded_for"),",") > 0 then
	str = mid(request.servervariables("http_x_forwarded_for"),1,instr(request.servervariables("http_x_forwarded_for"),",")-1)
  elseif instr(request.servervariables("http_x_forwarded_for"),";") > 0 then
	str = mid(request.servervariables("http_x_forwarded_for"),1,instr(request.servervariables("http_x_forwarded_for"),";")-1)
  else
	str = request.servervariables("http_x_forwarded_for")
  end if
  get_ip = trim(mid(str,1,30))
end function
'获取插件名称
function get_plu_name()
  dim plu:plu = ""
  if get_session("plugin","") = "" then
	dim url:url = get_url()
	dim len_:len_ = instrrev(url,".") - instr(url,".")
	if len_ > 0 then plu = mid(url,instr(url,".") + 1,len_ - 1)
  else
    plu = get_session("plugin","")
  end if
  get_plu_name = plu
end function

%>