<%
class footer_model
  private sit_domain
  private sit_code
  private sit_code_url
  private sit_tec
  private sit_tec_url
  private sit_count_code
  
  private sub class_initialize()
	call db.reset()
	call db.sql_query("select * from "&S_DB_PREFIX&"config where con_type = 'site'")
	if db.get_record_count() > 0 then
	  dim record_count:record_count = db.get_record_count()
	  dim i
	  for i = 0 to record_count - 1
		select case db.sql_result(i,"con_name")
		  case "sit_domain"
			sit_code = db.sql_result(i,"sit_domain")
		  case "sit_code"
			sit_code = db.sql_result(i,"con_value")
		  case "sit_code_url"
			sit_code_url = db.sql_result(i,"con_value")
		  case "sit_tec"
			sit_tec = db.sql_result(i,"con_value")
		  case "sit_tec_url"
			sit_tec_url = db.sql_result(i,"con_value")
		  case "sit_count_code"
			sit_count_code = db.sql_result(i,"con_value")
		end select
	  next
	end if
  end sub
	
  public function get_info(val)
	select case val
	  case "sit_domain"
	    get_info = sit_domain
	  case "sit_code"
	    get_info = sit_code
	  case "sit_code_url"
	    get_info = sit_code_url
	  case "sit_tec"
	    get_info = sit_tec
	  case "sit_tec_url"
	    get_info = sit_tec_url
	  case "sit_count_code"
	    get_info = sit_count_code
	end select
  end function
end class
'新秀
%>