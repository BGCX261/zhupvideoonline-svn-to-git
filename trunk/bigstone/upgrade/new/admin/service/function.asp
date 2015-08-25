<%
function get_mes_sheet(page,page_sum)
  dim i,arr()
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"message order by mes_id desc")
  dim record_count:record_count = db.get_record_count()
  if record_count > 0 then
	dim page_size:page_size = 10
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
	redim preserve arr(page_size - 1,5)
	for i = 0 to page_size - 1
	  arr(i,0) = db.sql_result(i,"mes_id")
	  arr(i,1) = db.sql_result(i,"mes_user")
	  arr(i,2) = db.sql_result(i,"mes_time")
	  arr(i,3) = db.sql_result(i,"mes_text")
	  arr(i,4) = db.sql_result(i,"mes_reply")
	  arr(i,5) = db.sql_result(i,"mes_show")
	next
  else
	redim preserve arr(-1)
	page_sum = 1
	page = 1
  end if
  get_mes_sheet = arr
end function

function get_res_sheet(page,page_sum)
  dim i,arr()
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"research order by res_id desc")
  dim record_count:record_count = db.get_record_count()
  if record_count > 0 then
	dim page_size:page_size = 10
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
	redim preserve arr(page_size - 1,3)
	for i = 0 to page_size - 1
	  arr(i,0) = db.sql_result(i,"res_id")
	  arr(i,1) = db.sql_result(i,"res_ip")
	  arr(i,2) = db.sql_result(i,"res_time")
	  arr(i,3) = db.sql_result(i,"res_text")
	  arr(i,3) = replace(replace(arr(i,3),"[]",""),"{v}","<br>")
	next
  else
	redim preserve arr(-1)
	page_sum = 1
	page = 1
  end if
  get_res_sheet = arr
end function

function get_link_sheet()
  dim i,record_count,arr()
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"link order by lin_index desc")
  record_count = db.get_record_count()
  if record_count > 0 then
	redim preserve arr(record_count - 1,5)
	for i = 0 to record_count - 1
	  arr(i,0) = db.sql_result(i,"lin_id")
	  arr(i,1) = db.sql_result(i,"lin_word")
	  arr(i,2) = db.sql_result(i,"lin_url")
	  arr(i,3) = db.sql_result(i,"lin_img")
	  arr(i,4) = db.sql_result(i,"lin_index")
	  arr(i,5) = db.sql_result(i,"lin_title")
	next
  else
	redim preserve arr(-1)
  end if
  get_link_sheet = arr
end function

function get_question()
  dim arr(),i,site,val
  call db.reset()
  call db.sql_query("select * from "&S_DB_PREFIX&"config where con_name = 'research'")
  if db.get_record_count() > 0 then
	dim record_count:record_count = db.get_record_count()
	redim preserve arr(record_count - 1,3)
	for i = 0 to record_count - 1
	  arr(i,0) = db.sql_result(i,"con_id")
	  val = db.sql_result(i,"con_value")
	  site = instr(val,"{v}")
	  arr(i,1) = left(val,site - 1)
	  val = right(val,len(val) - site - 2)
	  site = instr(val,"{v}")
	  arr(i,2) = left(val,site - 1)
	  arr(i,3) = replace(right(val,len(val) - site - 2),"{v}","|")
	next
  else
	redim preserve arr(-1)
  end if
  get_question = arr
end function

'新秀
%>