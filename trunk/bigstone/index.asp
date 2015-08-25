<!-- #include file="include/config.asp" -->
<%
dim url:url = request.servervariables("query_string")
if left(url,10) = "index.asp?" then url = right(url,len(url) - 10)
dim path
if S_STATIC > 0 then
  if url = "" then
	if S_FLASH <> 1 then
	  url = "index."&S_SUFFIX
	else
	  url = "flash."&S_SUFFIX
	end if
  end if
  path = S_ROOT&"html/"&url
  dim fso:set fso = server.createobject("scripting.filesystemobject")
  if fso.fileexists(server.mappath(path)) then
	if S_STATIC = 2 then
	  response.redirect(path)
	else
	  server.execute(path)
	end if
	response.end
  end if
  set fso = nothing
end if

if url = "" then
  if S_FLASH <> 1 then
	path = S_ROOT&"index/index.asp"
  else
	path = S_ROOT&"index/flash.asp"
  end if
else
  str = mid(url,1,instr(url,".") - 1)
  arr = split(str,"_")
  path = S_ROOT&"index/"&arr(0)&".asp"
end if
server.execute(path)
'新秀
%>
