<!-- #include file="../plugins/xheditor/common.asp" -->
<!-- #include file="../plugins/en/common.asp" -->
<%'include%>

<%
function interface(site,val)
select case site
case "reset_config":if val = "en" then call plu_en_reset()'reset_config
case "admin_editor":call plu_xheditor(val)'admin_editor
case "admin_header":call plu_en_select(val)'admin_header
end select
end function

'新秀
%>