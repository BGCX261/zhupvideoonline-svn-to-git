<%
class db_class
  private root
  private db_type
  private db_path
  private db_name
  private db_user
  private db_pwd
  private db_prefix
  private conn
  private rs
  private record_count
  private fields_count
  private page_size
  private fields_array
  private data_array
  private is_arr
  
  private sub class_initialize()
	record_count = -1
	fields_count = -1
	page_size = -1
	is_arr = 0
  end sub
  '连接数据库
  public function sql_connect()
    on error resume next
	select case db_type
	  case "access"
		dim str:str = "provider=microsoft.jet.oledb.4.0;data source="&server.mappath(root&db_path&db_name)
		set conn = server.createobject("adodb.connection")
		conn.open str
	end select
	if err then
	  err.clear
	  response.end
	end if
  end function
  '初始化
  public function set_db(root_,db_type_,db_path_,db_name_,db_user_,db_pwd_,db_prefix_)
	root = root_
	db_type = db_type_
	db_path = db_path_
	db_name = db_name_
	db_user = db_user_
	db_pwd = db_pwd_
	db_prefix = db_prefix_
  end function
  '关闭数据库
  public function close_db()
    rs.close
	set rs = nothing
	conn.close
	set conn = nothing
  end function
  '执行SQL语句
  public function sql_query(sql)
	select case db_type
	  case "access"
		select case sql_type(sql)
		  case "select"
			set rs = server.createobject("adodb.recordset")
			on error resume next
			rs.open sql,conn,1,3
			if err then
			  err.clear
			  response.write(sql)
			  response.end
			end if
			if not rs.eof then
			  record_count = rs.recordcount
			  fields_count = rs.fields.count
			else
			  record_count = 0
			end if
		  case "insert"
			conn.execute(sql)
		  case "update"
			conn.execute(sql)
		  case "delete"
		    conn.execute(sql)
		end select
	end select
  end function
  '获取结果数据
  public function sql_result(row,field_name)
    dim i
	if is_arr = 0 then
	  call sql_result_array()
	  is_arr = 1
	end if
	for i = 0 to fields_count - 1
	  if fields_array(i) = field_name then
		sql_result = data_array(row * fields_count + i)
	  end if
	next
  end function
  '获取记录数
  public function get_record_count()
	get_record_count = record_count
  end function
  '获取字段数
  public function get_fields_count()
	get_fields_count = fields_count
  end function
  '设置每页记录数
  public function set_page_size(val)
  	select case db_type
	  case "access"
		rs.pagesize = int(val)
	end select
  end function
  '设置当前页
  public function set_absolute_page(val)
  	select case db_type
	  case "access"
		rs.absolutepage = val
	end select
  end function
  '设置当前页实际记录数
  public function set_real_page_size(val)
	page_size = val
  end function
  '重置
  public function reset()
	record_count = -1
	fields_count = -1
	page_size = -1
	is_arr = 0
  end function
  '判断查询类型
  private function sql_type(sql)
	select case left(sql,6)
	  case "select"
		sql_type = "select"
	  case "insert"
		sql_type = "insert"
	  case "update"
		sql_type = "update"
	  case "delete"
		sql_type = "delete"
	end select
  end function
  '将记录集转换为数组
  private function sql_result_array()
	dim i,j
	dim str:str = ""
	for i = 0 to fields_count - 1
	  str = str & "{^}" & replace(rs.fields(i).name&"","{^}","")
	next
	str = mid(str,4,len(str))
	fields_array = split(str,"{^}")
	str = ""
	dim num
	if page_size <> -1 then
	  num = page_size
	else
	  num = record_count
	end if
	for i = 0 to num - 1
	  for j = 0 to fields_count - 1
		str = str & "{^}" & replace(rs(rs.fields(j).name)&"","{^}","")
	  next
	  rs.movenext
	next
	str = mid(str,4,len(str))
	data_array = split(str,"{^}")
  end function
end class
'新秀
%>