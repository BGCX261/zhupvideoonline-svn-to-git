<%if get_session("login","") = "" or S_ROOT = "" then response.end%>
<div class="box" id="result">
  <div class="close" onclick="hid_box('result')">关闭</div>
  <div class="main">执行完毕</div>
</div>
<div class="box" id="del">
  <div class="close" onclick="hid_box('del')">关闭</div>
  <div class="main">您确定要删除该信息吗？
    <div class="button">
      <form name="form_del" method="post">
        <input name="id" id="del_id" type="hidden" value="" />
        <input type="button" onclick="ajax_submit('form_del')" value="确定" />
        <input type="button" onclick="hid_box('del')" value="取消" />
      </form>
    </div>
  </div>
</div>
<div class="box" id="no_deal">
  <div class="close" onclick="hid_box('no_deal')">关闭</div>
  <div class="main">该分类存在子类，请先操作所有子类</div>
</div>