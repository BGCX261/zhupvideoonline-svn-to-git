<%
'模板引擎
class sinty_
  private arr_a()
  private arr_b()
  private arr_len
  private template
  private file_name
  private tpl_path
  private sinty_path
  private include
  private code_page
  private compile
  private rewrite
  private cache
  private reg
  
  private sub class_initialize()
	redim arr_a(0)
	redim arr_b(0)
	arr_len = 0
	code_page = 65001
	include = ""
	compile = 1
	rewrite = "?"
	set reg = new regexp
	reg.ignorecase = true
	reg.global = true
  end sub
  
  public function get_tpl_info(val)
	tpl_path = val
	template = get_template(tpl_path)
	file_name = get_name(tpl_path,"/")
	dim sinty_xml:sinty_xml = server.mappath(left(tpl_path,len(tpl_path)-len(file_name))&"sinty.xml")
	file_name = left(file_name,len(file_name)-len(get_name(file_name,"."))-1)
	on error resume next
	dim xml:set xml = server.createobject("microsoft.xmldom")
	xml.load(sinty_xml)
	compile = xml.selectsinglenode("site").getattribute("compile")
	if compile <> 2 then
	  page_compile = xml.selectsinglenode("site").selectsinglenode(file_name).getattribute("compile")
	  if compile <> page_compile then compile = page_compile
	end if
	dim file_xml:set file_xml = xml.selectsinglenode("site").selectsinglenode(file_name)
	dim i,text
	dim nodes_num:nodes_num = file_xml.childnodes.length
	dim return:return = ""
	for i = 0 to nodes_num - 1
	  text = file_xml.childnodes.item(i).text
	  if left(text,1) <> "*" then return = return & text & "|"
	next
	set file_xml = nothing
	set xml = nothing
	if err then
	  err.clear
	  get_tpl_info = ""
	else
	  get_tpl_info = left(return,len(return)-1)
	end if
  end function
  
  public function assign(val_a,val_b)
	redim preserve arr_a(arr_len)
	arr_a(arr_len) = val_a
	redim preserve arr_b(arr_len)
	arr_b(arr_len) = val_b
	arr_len = arr_len + 1
  end function
  
  public function display()
	call set_data()
	if compile = 1 then
	  dim fso:set fso = server.createobject("scripting.filesystemobject")
	  if not fso.fileexists(server.mappath(sinty_path&"compile/"&file_name&".asp")) then
		call do_compile()
	  end if
	  set fso = nothing
	elseif compile = 2 then
	  call do_compile()
	end if
	set reg = nothing
	server.execute(sinty_path&"cache/"&cache&".asp")
  end function
  
  public function found_html(path)
	dim url:url = "http://"&request.servervariables("server_name")&request("http")&sinty_path&"cache/"&cache&".asp"
	dim xml_http:set xml_http = server.createobject("microsoft.xmlhttp")
	xml_http.open "get",url,false
	xml_http.send()
	dim text:text = xml_http.responsebody
	dim stm:set stm = server.createobject("adodb.stream")
	with stm
	  .type = 1
	  .open()
	  .write text
	  .savetofile path,2
	  .close()
	end with
	set xml_http = nothing
	set stm = nothing
  end function

  public function set_sinty_path(val)
	sinty_path = val
  end function
  
  public function set_include(val)
	include = include&"|"&val
  end function  
  
  public function set_code_page(val)
	code_page = val
  end function
  
  public function set_rewrite(val)
	rewrite = val
  end function
    
  private function get_template(val)
    dim path:path = server.mappath(val)
	dim stm:set stm = server.createobject("adodb.stream")
	with stm
	  .type = 2
	  .mode = 3
	  .charset = "utf-8"
	  .open
	  .loadfromfile path
	  dim tpl:tpl = .readtext
	  .close
	end with
	set stm = nothing
	get_template = tpl
  end function
  
  function do_compile()
    call replace_inc()
	call replace_asp_end()
	call replace_note()
	call replace_url()
	call replace_simple()
	reg.pattern = "\{.*?\}"
	dim mat
	dim matches:set matches = reg.execute(template)
	for each mat in matches
	  dim tag:tag = mat.value
	  if left(tag,4) = "{if:" then
		call replace_if(mid(tag,5,len(tag)-5))
	  elseif left(tag,8) = "{elseif:" then
		call replace_elseif(mid(tag,9,len(tag)-9))
	  elseif left(tag,8) = "{select:" then
		call replace_select(mid(tag,9,len(tag)-9))
	  elseif left(tag,6) = "{case:" then
		call replace_case(mid(tag,7,len(tag)-7))
	  elseif left(tag,5) = "{for:" then
		call replace_for(mid(tag,6,len(tag)-6))
	  elseif left(tag,6) = "{func:" then
		call replace_func(mid(tag,7,len(tag)-7))
	  end if
	next
	call replace_arr()
	call replace_var()
	reg.pattern = "\{[(if)|(elseif)|(select)|(for)|(call)]/S.*?\}"
	dim val
	set matches = reg.execute(template)
	for each mat in matches
	  val = replace(mat.value,"{",chr(60)&chr(37))
	  val = replace(val,"}",chr(37)&chr(62))
	  template = replace(template,mat.value,val)
	next
	set matches = nothing
	dim path:path = server.mappath(sinty_path&"compile/"&file_name&".asp")
	call put_out(path,template)
  end function
  
  private function replace_inc()
	do while instr(template,"{inc:") > 0
	  reg.pattern = "\{inc:.*?\}"
	  dim mat
	  dim matches:set matches = reg.execute(template)
	  for each mat in matches
		dim val:val = mid(mat.value,6,len(mat.value) - 6)
		dim inc_file:inc_file = left(tpl_path,len(tpl_path)-len(get_name(tpl_path,"/")))&val&"."&get_name(tpl_path,".")
		template = replace(template,"{inc:"&val&"}",get_template(inc_file))  	  
	  next
	  set matches = nothing
	loop  
  end function
  
  private function replace_asp_end()
	template = replace(template,chr(60)&chr(37)&"response.end"&chr(37)&chr(62),"")
  end function
  
  private function replace_note()
	do while instr(template,"{*") > 0
	  reg.pattern = "\{\*.*?\*\}"
	  dim mat
	  dim matches:set matches = reg.execute(template)
	  for each mat in matches
		dim val:val = mid(mat.value,3,len(mat.value) - 4)
		template = replace(template,"{*"&val&"*}","")
	  next
	  set matches = nothing
	loop
  end function
  
  private function replace_url()
	if rewrite <> "?" then template = replace(template,"{$S_ROOT}?","{$S_ROOT}"&rewrite)
  end function
  
  private function replace_simple()
	template = replace(template,"{else}",chr(60)&chr(37)&"else"&chr(37)&chr(62))
	template = replace(template,"{/if}",chr(60)&chr(37)&"end if"&chr(37)&chr(62))
	template = replace(template,"{/for}",chr(60)&chr(37)&"next"&chr(37)&chr(62))
	template = replace(template,"{/select}",chr(60)&chr(37)&"end select"&chr(37)&chr(62))
	template = replace(template,"{default}",chr(60)&chr(37)&"case else"&chr(37)&chr(62))
  end function

  private function replace_arr()
	reg.pattern = "\{.*?\}"
	dim matches:set matches = reg.execute(template)
	dim mat
	for each mat in matches
	  dim val:val = mat.value
	  dim var,matches_b,mat_b,arr_name,myarr
	  dim i,index_a,index_b,myindex
	  reg.pattern = "\$[^$|^\{|^,|^\]]*?\["
	  set matches_b = reg.execute(val)
	  for each mat_b in matches_b
		arr_name = replace(mat_b.value,"$","")
		arr_name = replace(arr_name,"[","")
	  next
	  reg.pattern = "\]\[\$.*?\]"
	  set matches_b = reg.execute(val)
	  for each mat_b in matches_b
		var = replace(mat_b.value,"][$","")
		var = replace(var,"]","")
		for i = 0 to ubound(arr_a)
		  if arr_a(i) = var then
			index_a = i
			myindex = i
		  end if
		next
		if typename(arr_b(index_a)) = "String" then
		  for i = 0 to ubound(arr_a)
			if arr_a(i) = arr_name&"_name" then index_b = i
		  next
		  myarr = arr_b(index_b)
		  for i = 0 to ubound(myarr)
			if myarr(i) = arr_b(index_a) then myindex = i
		  next
		end if
		val = replace(val,"[$"&var&"]","["&myindex&"]")
	  next
	  reg.pattern = "\]\[.*?\]"
	  set matches_b = reg.execute(val)
	  for each mat_b in matches_b
		var = replace(replace(replace(mat_b.value,"][",""),"]",""),"'","")
		index_a = -1
		index_b = 0
		for i = 0 to ubound(arr_a)
		  if arr_a(i) = arr_name&"_name" then index_a = i
		next
		if index_a <> -1 then
		  myarr = arr_b(index_a)
		  for i = 0 to ubound(myarr)
			if myarr(i) = var then index_b = i
		  next
		end if
		val = replace(mat.value,"['"&var&"']","["&index_b&"]")
	  next
	  val = replace(val,"][",",")
	  val = replace(val,"[","(")
	  val = replace(val,"]",")")
	  template = replace(template,mat.value,val)
	next
	set matches = nothing
  end function
  
  private function replace_var()
	reg.pattern = "\{[^\{]*?\$[a-z]{1}[^\n|^;]*?\}"
	dim matches:set matches = reg.execute(template)  
	dim mat
	for each mat in matches
	  dim val:val = mat.value
	  if instr(val,"=") > 0 then val = replace(val,"{$",chr(60)&chr(37)) else val = replace(val,"{$",chr(60)&chr(37)&"=")
	  val = replace(val,"$","")
	  val = replace(val,"{",chr(60)&chr(37))
	  val = replace(val,"}",chr(37)&chr(62))	  
	  template = replace(template,mat.value,val)
	next
	set matches = nothing
  end function
  
  private function replace_if(val)
    dim arr:arr = split(val,",")
	template = replace(template,"{if:"&val&"}","{if "&arr(0)&get_mark(arr(1))&arr(2)&" then}")
  end function
  
  private function replace_elseif(val)
    dim arr:arr = split(val,",")
	template = replace(template,"{elseif:"&val&"}","{elseif "&arr(0)&get_mark(arr(1))&arr(2)&" then}")
  end function
  
  private function replace_select(val)
	template = replace(template,"{select:"&val&"}","{select case "&val&"}")
  end function
  
  private function replace_case(val)
	template = replace(template,"{case:"&val&"}",chr(60)&chr(37)&"case "&chr(34)&val&chr(34)&":"&chr(37)&chr(62))
  end function
  
  private function replace_for(val)
	dim arr:arr = split(val,",")
	template = replace(template,"{for:"&val&"}","{for "&arr(0)&"="&arr(1)&" to "&arr(2)&"}")
  end function
  
  private function replace_func(val)
	dim arr:arr = split(val,",")
	dim str:str = ""
	dim i
	for i = 1 to ubound(arr)
	  str = str & "," & arr(i)
	next
	str = right(str,len(str) - 1)
	template = replace(template,"{func:"&val&"}","{call "&arr(0)&"("&str&")}")
  end function
  
  private function set_inc()
	dim str:str = ""
	dim br:br = chr(13)&chr(10)
	if include <> "" then
	  include = right(include,len(include)-1)
	  dim arr_inc:arr_inc = split(include,"|")
	  dim i
	  for i = 0 to ubound(arr_inc)
		str = str & chr(60)&"!-- #include file="&chr(34)&arr_inc(i)&chr(34)&" --"&chr(62)&br
	  next
	end if
	set_inc = str
  end function
  
  private function get_mark(val)
	select case val
	  case "eq":get_mark = "="
	  case "neq":get_mark = "<>"
	  case "gt":get_mark = ">"
	  case "egt":get_mark = ">="
	  case "lt":get_mark = "<"
	  case "elt":get_mark = "<="
	end select
  end function
  
  private function filter_code(str)
	str = replace(str,chr(60)&chr(37),"")
	str = replace(str,chr(37)&chr(62),"")
	str = replace(str,chr(13)&chr(10),"")
	str = replace(str,chr(34),chr(34)&chr(34))
	str = replace(str,"&$39;",chr(39))
	filter_code = str
  end function
  
  private function get_name(file_path,str)
	if file_path <> "" then
	  get_name = mid(file_path,instrrev(file_path,str) + 1)
	else
	  get_name = ""
	end if
  end function
  
  private function set_data()
	dim br:br = chr(13)&chr(10)
	if arr_len > 0 or file_name = file_name then
	  dim data:data = data & chr(60)&chr(37)&br
	  dim i,j,k
	  for i = 0 to ubound(arr_a)
		dim type_name:type_name = typename(arr_b(i))
		if type_name = "String" then
		  data = data & arr_a(i)&" = "&chr(34)&filter_code(arr_b(i))&chr(34)&br
		elseif type_name = "Variant()" then
		  on error resume next
			dim try:try = ubound(arr_b(i),2)
			dim arr_type:arr_type = 2
		  if err then
			err.clear
			arr_type = 1
		  end if
		  if arr_type = 1 then
			dim arr:arr = arr_b(i)
			data = data & "dim "&arr_a(i)&"("&ubound(arr)&")"&br
			for j = 0 to ubound(arr)
			  data = data & arr_a(i)&"("&j&") = "&chr(34)&filter_code(arr(j))&chr(34)&br
			next	
		  elseif arr_type = 2 then
			dim len_a:len_a = ubound(arr_b(i),1)
			dim len_b:len_b = ubound(arr_b(i),2)
			arr = arr_b(i)
			data = data & "dim "&arr_a(i)&"("&len_a&","&len_b&")"&br
			for j = 0 to len_a
			  for k = 0 to len_b
				data = data & arr_a(i)&"("&j&","&k&") = "&chr(34)&filter_code(arr(j,k))&chr(34)&br
			  next
			next		  	
		  end if
		elseif type_name = "Integer" or type_name = "Long" or type_name = "Single" or type_name = "Double" then
		  data = data & arr_a(i)&" = "&arr_b(i)&br
		end if
	  next
	  data = data & chr(37)&chr(62)&br
	end if
	data = set_inc() & data & chr(60)&"!-- #include file="&chr(34)&"../compile/"&file_name&".asp"&chr(34)&" --"&chr(62)&br
	cache = file_name & "_" & (second(time()) mod 5)
	dim path:path = server.mappath(sinty_path&"cache/"&cache&".asp")
	call put_out(path,data)
  end function
  
  private function put_out(path,text)
	dim stm:set stm = server.createobject("adodb.stream")
	with stm
	  .type = 2
	  .mode = 3
	  .charset = "utf-8"
	  .open
	  .writetext text
	  .savetofile path,2
	  .close
	end with
	set stm = nothing
  end function
  
end class
%>