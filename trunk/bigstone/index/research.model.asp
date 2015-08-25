<%
class research_model
  private research()
  private research_count
  
  private sub class_initialize()
	
  end sub
	
  public function get_research()
	call db.reset()
	call db.sql_query("select * from "&S_DB_PREFIX&"config where con_name = 'research'")
	if db.get_record_count() > 0 then
	  dim record_count:record_count = db.get_record_count()
	  research_count = record_count
	  redim preserve research(record_count - 1,1)
	  dim i
	  for i = 0 to record_count - 1
		research(i,0) = db.sql_result(i,"con_id")
		research(i,1) = db.sql_result(i,"con_value")
	  next
	else
	  record_count = 0
	end if
	get_research = research
  end function
  
  public function get_research_count()
	get_research_count = research_count
  end function
  public function get_research_name()
	dim research_name(1)
	research_name(0) = "id"
	research_name(1) = "text"	
	get_research_name = research_name
  end function
	
end class
'新秀
%>