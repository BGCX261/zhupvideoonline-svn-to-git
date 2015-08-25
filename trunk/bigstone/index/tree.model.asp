<%
class tree_model
  private tree()
  private tree_count
  
  private sub class_initialize()
	
  end sub
	
  public function get_tree(cat_type)
    dim i
	dim flag:flag = 0
	call db.reset()
	call db.sql_query(get_sql(cat_type))
	dim record_count:record_count = db.get_record_count()
	tree_count = record_count
	if record_count > 0 then
	  redim preserve tree(record_count - 1,2)
	  for i = 0 to record_count - 1
		tree(i,0) = db.sql_result(i,"cat_id")
		tree(i,1) = db.sql_result(i,"cat_parent_id")
		tree(i,2) = db.sql_result(i,"cat_name")
		if tree(i,1) <> 0 then flag = 1
	  next
	  get_tree = set_cat_order(tree,0,1,flag)
	else
	  tree_count = 0
	  get_tree = tree
	end if
  end function

  private function get_sql(cat_type)
	get_sql = "select * from "&S_DB_PREFIX&"category where cat_type = '"&cat_type&"' and cat_show = 1 order by cat_top desc,cat_index desc,cat_id asc"
  end function
  
  '对分类进行排序，调用了get_tree_order()，并对其返回结果进行处理，a代表ID下标，b代表父类ID下标，flag标志是否存在第二级分类
  private function set_cat_order(arr,a,b,flag)
	dim i,j,k,tree_order,arr1(),arr2(),arr3,arr4,arr5()
	dim u1:u1 = ubound(arr)
	dim u2:u2 = ubound(arr,2)
	if u1 > -1 and flag = 1 then
	  redim preserve arr1(u1)
	  redim preserve arr2(u1)
	  for i = 0 to u1
		arr1(i) = arr(i,a)
		arr2(i) = arr(i,b)
	  next
	  if u1 > -1 then
		tree_order = split(get_tree_order(arr1,arr2),"{v}")
		arr3 = split(tree_order(0),"|")
		arr4 = split(tree_order(1),"|")
		redim preserve arr5(u1,u2 + 1)
		for i = 1 to ubound(arr3) - 1
		  for j = 0 to u1
			if arr(j,0) = arr3(i) then
			  for k = 0 to u2
				arr5(i - 1,k) = arr(j,k)
			  next
			  arr5(i - 1,u2 + 1) = arr4(i)
			end if
		  next
		next
	  end if
	elseif u1 > -1 and flag = 0 then
	  redim preserve arr5(u1,u2 + 1)
	  for i = 0 to u1
		for k = 0 to u2
		  arr5(i,k) = arr(i,k)
		next
		arr5(i,u2 + 1) = 1
	  next
	else
	  redim preserve arr5(-1)
	end if
	set_cat_order = arr5
  end function
  
  '对无序多级分类进行排序，返回一个记录排序信息的特殊字符串
  private function get_tree_order(id,parent)
	dim i,id_arr,parent_arr,cat_str
	dim j:j = 0
	dim k:k = ubound(parent) + 1
	dim id_str()
	dim parent_str()
	dim grade_str:grade_str = "|"
	'将分类划分为不同等级
	do while k > 0
	  redim preserve id_str(j)
	  redim preserve parent_str(j)
	  id_str(j) = "|"
	  parent_str(j) = "|"
	  for i = 0 to ubound(parent)
		if j = 0 then
		  if parent(i) = 0 then
			id_str(j) = id_str(j) & id(i) & "|"
			parent_str(j) = parent_str(j) & parent(i) & "|"
			k = k - 1
		  end if
		else
		  if instr(id_str(j - 1),"|"&parent(i)&"|") > 0 then
			id_str(j) = id_str(j) & id(i) & "|"
			parent_str(j) = parent_str(j) & parent(i) & "|"
			k = k - 1
		  end if
		end if
	  next
	  j = j + 1
	loop
	'将子级分类倒序排列
	for i = 1 to ubound(id_str)
	  dim str1:str1 = ""
	  dim str2:str2 = ""
	  id_arr = split(id_str(i),"|")
	  parent_arr = split(parent_str(i),"|")
	  for j = 1 to ubound(id_arr)
		str1 = id_arr(j) & "|" & str1
		str2 = parent_arr(j) & "|" & str2
	  next
	  id_str(i) = str1
	  parent_str(i) = str2
	next
	'将子级分类插入父级分类后面
	if ubound(id_str) > 0 then cat_str = id_str(0)
	for i = 1 to ubound(id_str)
	  id_arr = split(id_str(i),"|")
	  parent_arr = split(parent_str(i),"|")
	  for j = 1 to ubound(parent_arr) - 1
		if instr(cat_str,"|"&parent_arr(j)&"|") > 0 then
		  cat_str = replace(cat_str,"|"&parent_arr(j)&"|","|"&parent_arr(j)&"|"&id_arr(j)&"|")
		end if
	  next
	next
	'获取等级信息
	dim arr:arr = split(cat_str,"|")
	for i = 1 to ubound(arr) - 1
	 for j = 0 to ubound(id_str)
	   if instr(id_str(j),"|"&arr(i)&"|") > 0 then grade_str = grade_str & (j + 1) & "|"
	 next
	next
	get_tree_order = cat_str & "{v}" & grade_str
  end function
  
  public function get_tree_count()
	get_tree_count = tree_count
  end function
  
  public function get_tree_name()
	dim tree_name(3)
	tree_name(0) = "id"
	tree_name(1) = "parent"
	tree_name(2) = "name"
	tree_name(3) = "grade"
	get_tree_name = tree_name
  end function
	
end class
'新秀
%>