<!-- #include file="../include/common.asp" -->
<!-- #include file="sinty/sinty.asp" -->
<!-- #include file="module.func.asp" -->

<%
dim sinty:set sinty = new sinty_
dim info:info = initial("about.asp")
dim url:url = get_url()
dim cat,page,id,plu
call set_parameter(0,0,0)
call set_head("article")
call select_module(split(info,"|"))
call sinty.display()
call found_html("about."&S_SUFFIX)
call db.close_db()
set db = nothing
set sinty = nothing

'新秀
%>