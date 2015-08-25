<!-- #include file="../system/common.asp" -->
<!-- #include file="../system/head.asp" -->
<!-- #include file="function.asp" -->
<%
dim page:page = gets("page","")
dim page_sum:page_sum = 0
dim arr,i
%>
<div class="here">问卷调查</div>
<div class="block">
  <form name="form_edit_question" method="post" action="deal.asp" target="deal">
    <input name="cmd" type="hidden" value="edit_question" />
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><b>问题</b></td>
        <td><b>表单类型</b></td>
        <td><b>答案选项</b></td>
        <td><b>操作</b></td>
      </tr>
      <%arr = get_question()%>
      <%for i = 0 to ubound(arr)%>
      <tr>
        <td><input name="question" type="text" class="text" maxlength="50" value="<%=arr(i,1)%>" /></td>
        <td>
        <select name="input_type" >
          <option value="radio" <%if arr(i,2) = "radio" then echo "selected"%>>单选框</option>
          <option value="checkbox" <%if arr(i,2) = "checkbox" then echo "selected"%>>复选框</option>
          <option value="text" <%if arr(i,2) = "text" then echo "selected"%>>文本框</option>
        </select>
        </td>
        <td><input name="answer" type="text" class="text" maxlength="150" value="<%=arr(i,3)%>" /></td>
        <td><input name="id" type="hidden" value="<%=arr(i,0)%>" /><span class="red" onClick="del('question,<%=arr(i,0)%>')">[删除]</span></td>
      </tr>
      <%next%>
      <tr>
        <td class="button" colspan="30">
          <input type="button" onClick="do_submit('form_edit_question')" value="修改" />
          <input type="button" onClick="show_box('add_question',350,175)" value="添加" />
        </td>
      </tr>
    </table>
  </form>
</div>
<div class="here">调查结果</div>
<div class="block">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><b>访客提交问卷</b></td>
      <td width="120px"><b>操作</b></td>
    </tr>
    <%arr = get_res_sheet(page,page_sum)%>
    <%for i = 0 to ubound(arr)%>
    <tr>
      <td style="text-align:left;line-height:20px;">
        <div><b>IP <%=arr(i,1)%> 的访客</b>&nbsp;<%=arr(i,2)%>&nbsp;提交问卷：</div>
        <div><%=arr(i,3)%></div>
      </td>
      <td>
        <a onClick="del('research,<%=arr(i,0)%>')">[删除]</a>
      </td>
    </tr>
    <%next%>
    <tr>
      <td colspan="30"><%call page_link("research.asp",page,page_sum,"","")%></td>
    </tr>
  </table>
</div>
<div class="here">使用说明</div>
<div class="block content">
  1.答案选项之间使用|号隔开，个数不限，例如“答案一|答案二|答案三”。
</div>
<iframe id="deal" name="deal"></iframe>
<!-------------------------- BOX -------------------------->
<div class="box" id="add_question">
  <div class="head">
    <div class="title">添加问题</div>
    <div class="close" onclick="hid_box('add_question')">关闭</div>
  </div>
  <form name="form_add_question" method="post">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>问题：</td>
        <td class="f_l"><input name="question" type="text" class="text" maxlength="50" /></td>
      </tr>
      <tr>
        <td>表单类型：</td>
        <td class="f_l">
        <select name="input_type" >
          <option value="radio">单选框</option>
          <option value="checkbox">复选框</option>
          <option value="text">文本框</option>
        </select>
        </td>
      </tr>
      <tr>
        <td>答案选项：</td>
        <td class="f_l"><input name="answer" type="text" class="text" maxlength="150" /></td>
      </tr>
      <tr>
        <td class="button" colspan="30"><input type="button" onclick="ajax_submit('form_add_question')" value="提交" /></td>
      </tr>
    </table>
  </form>
</div>
<!-------------------------- BOX -------------------------->
<script language="javascript">
var page = "<%=page%>";
interval = setInterval("get_result()",1000);
function do_submit(val)
{
	switch(val)
	{
		case "form_edit_question":document.form_edit_question.submit();break;
	}
	interval = setInterval("get_result()",1000);
}
function ajax_submit(val)
{
	switch(val)
	{
		case "form_add_question":
		var question = document.form_add_question.question.value;
		var input_type = document.form_add_question.input_type.value;
		var answer = document.form_add_question.answer.value;
		ajax("post","deal.asp","cmd=add_question&question="+question+"&input_type="+input_type+"&answer="+answer,
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
		case "form_del":
		var id = document.form_del.id.value;
		var arr = id.split(",");
		ajax("post","deal.asp","cmd=del_"+arr[0]+"&id="+arr[1],
		function(data)
		{
			if(data == 1) location.replace(location.href);
		});
		break;
	}
}
</script>
</body>
</html>
<%
call db.close_db()
set db = nothing
'新秀
%>