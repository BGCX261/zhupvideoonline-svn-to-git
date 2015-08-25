<!-- #include file="../include/common.asp" -->
<!-- #include file="sinty/sinty.asp" -->
<!-- #include file="module.func.asp" -->

<%
dim sinty:set sinty = new sinty_
dim info:info = initial("search.asp")
dim url:url = get_url()
dim cat,page,id,plu
dim keywords:keywords = post("search","")
if keywords <> "" then response.cookies("search") = keywords else keywords = strict(request.cookies("search"))
call set_parameter(0,0,0)
call set_head("search")
call reset_head()
call select_module(split(info,"|"))
call sinty.display()
call db.close_db()
set db = nothing
set sinty = nothing

function reset_head()
  call sinty.assign("page_title",keywords)
end function

'新秀
%>