<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%dim arr,i%>
<div class="here">当前使用模板文件</div>
<div class="block sheet">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><b>文件名</b></td>
      <td width="150px"><b>操作</b></td>
    </tr>
    <%arr = get_tpl_files()%>
    <%for i = 0 to ubound(arr)%>
    <tr>
      <td><%=arr(i)%></td>
      <td>
        <a href="tpl_edit.asp?fil=<%=arr(i)%>">[编辑]</a>
      </td>
    </tr>
    <%next%>
    <tr>
      <td colspan="30" height="30px"></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
call db.close_db()
set db = nothing
'新秀
%>