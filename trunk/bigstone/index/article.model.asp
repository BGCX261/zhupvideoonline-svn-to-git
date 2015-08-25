<%
class article_model
  private article()
  private fields_
  private fie_arr
  private art_id
  private cat_id
  private top
  private art_type
  private order_by
  private page_size
  private page
  private page_sum
  private record_sum
  private family
  
  private sub class_initialize()
	fields_ = ""
    art_id = 0
    cat_id = 0
	top = 0
	art_type = "default"
	order_by = " order by art_top desc,art_index desc,art_id desc"
	page_size = 10
	page = 1
	page_sum = 1
	record_sum = 1
  end sub
  
  public function get_article()
	call db.reset()
	call db.sql_query(get_sql("="))
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
	  redim preserve article(page_size - 1,arr_u)
	  dim i,j
	  for i = 0 to page_size - 1
		for j = 0 to arr_u
		  if left(fie_arr(j),3) = "art" then
			article(i,j) = db.sql_result(i,fie_arr(j))
		  else
			article(i,j) = replace_att(get_attribute(db.sql_result(i,"art_attributes"),fie_arr(j)))
		  end if
		next
	  next
	else
	  page_size = 0
	end if
	get_article = article
  end function
  
  public function get_prev(val)
	order_by = " order by art_top asc,art_index asc,art_id asc"
	call db.reset()
	call db.sql_query(get_sql(">"))
	get_prev = db.sql_result(0,"art_"&val)
  end function
  
  public function get_next(val)
	order_by = " order by art_top desc,art_index desc,art_id desc"
	call db.reset()
	call db.sql_query(get_sql("<"))
	get_next = db.sql_result(0,"art_"&val)
  end function

  private function get_sql(mark)
    dim table:table = S_DB_PREFIX&"article"
	dim where:where = " where art_show = 1 and art_type = '"&art_type&"' "
	if art_id <> 0 then
	  where = where & "and art_id "&mark&" "&art_id
	elseif cat_id <> 0 then
	  dim cat_str:cat_str = set_cat_str(cat_id)
	  table = table&","&S_DB_PREFIX&"category"
	  where = where & "and ("&cat_str&") and art_cat_id = cat_id"
	end if
	if top = 1 then where = where&"and art_top = 1 "
	get_sql = "select "&fields_&" from "&table&where&order_by
  end function
  
  private function set_cat_str(cat_id)
	dim i,str:str = ""
	if typename(family) = "Variant()" then
	  for i = 0 to ubound(family)
		str = str & "art_cat_id = "&family(i)&" or "
	  next
	  str = left(str,len(str)-4)
	else
	  str = "art_cat_id = " & cat_id
	end if
	set_cat_str = str
  end function
  
  private function replace_att(str)
	str = replace(str,"http;//","http://")
	replace_att = str
  end function
  
  public function set_fields(val)
    fie_arr = split(val,",")
	fields_ = ""
	dim flag:flag = 0
	dim i:for i = 0 to ubound(fie_arr)
	  if left(fie_arr(i),3) = "art" then
		fields_ = fields_ & fie_arr(i) &","
	  else 
	    flag = 1
	  end if 
	next
	if flag = 0 then
	  fields_ = left(fields_,len(fields_) - 1)
	else
	  fields_ = fields_ & "art_attributes"
	end if
  end function
  
  public function set_art_id(val)
	art_id = val
  end function
  
  public function set_cat_id(val)
	cat_id = val
  end function
  
  public function set_family(val)
	family = val
  end function
  
  public function set_art_type(val)
	art_type = val
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
  
  public function set_top(val)
	top = val
  end function
  
  public function get_article_name()
	get_article_name = fie_arr
  end function

end class
'新秀
%>