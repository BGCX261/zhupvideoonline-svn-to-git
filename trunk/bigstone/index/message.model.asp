<%
class message_model
  private message()
  private fields_
  private fie_arr
  private mes_id
  private order_by
  private page_size
  private page
  private page_sum
  private record_sum
  
  private sub class_initialize()
	fields_ = ""
    mes_id = 0
	order_by = " order by mes_id desc"
	page_size = 10
	page = 1
	page_sum = 1
	record_sum = 1
  end sub
  
  public function get_message()
	call db.reset()
	call db.sql_query(get_sql())
	if db.get_record_count() > 0 then
	  dim record_count:record_count = db.get_record_count()
	  record_sum = record_count
	  call db.set_page_size(page_size)
	  if (record_count mod page_size) > 0 then
		page_sum = int(record_count / page_size) + 1
	  else
		page_sum = int(record_count / page_size)
	  end if
	  page = num_bound(1,page_sum,page)
	  call db.set_absolute_page(page)
	  record_count = record_count - (page - 1) * page_size
	  if record_count < page_size then
		page_size = record_count
	  end if
	  call db.set_real_page_size(page_size)
	  dim arr_u:arr_u = ubound(fie_arr)
	  redim preserve message(page_size - 1,arr_u)
	  dim i,j
	  for i = 0 to page_size - 1
		for j = 0 to arr_u
		  message(i,j) = db.sql_result(i,fie_arr(j))
		next
	  next
	else
	  page_size = 0
	end if
	get_message = message
  end function
  
  private function get_sql()
	get_sql = "select "&fields_&" from "&S_DB_PREFIX&"message where mes_show = 1"&order_by
  end function
  
  public function set_fields(val)
	fields_ = val
	fie_arr = split(val,",")
  end function
  
  public function set_mes_id(val)
	art_id = val
  end function
  
  public function set_order_by(val)
	if val <> "" then order_by = val
  end function
  
  public function set_page_size(val)
	page_size = int(val)
  end function
  
  public function set_page_num(val)
	page = val
  end function
  
  public function get_page_size()
	get_page_size = page_size
  end function
  
  public function get_page_sum()
	get_page_sum = page_sum
  end function
  
  public function get_record_sum()
	get_record_sum = record_sum
  end function
  
  public function get_message_name()
	get_message_name = fie_arr
  end function

end class
'新秀
%>