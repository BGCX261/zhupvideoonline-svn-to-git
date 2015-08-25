<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<%
dim dir:dir = gets("dir","")
dim fil:fil = gets("fil","")
dim extension:extension = lcase(get_file_name(fil,"."))
if gets("reduce","") = "1" then
  if instr("jpg,gif,png,bmp,jpeg",extension) > 0 then
	call upload(dir,fil)
	call reduce(dir,fil,int(get_config("reduce_img_width")),int(get_config("reduce_img_height")))
	call set_session("img",dir&fil,"")
  end if
else
  if instr("asp,php,jsp,aspx",extension) = 0 then
	call upload(dir,fil)
	call set_session("fil",dir&fil,"")
  end if
end if
%>
</body>
</html>
<%
call db.close_db()
set db = nothing
'新秀
%>