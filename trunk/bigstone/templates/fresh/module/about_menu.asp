<%response.end%>
<div class="box tree">
  <div class="head"><span>{$L_ABOUT_MENU}</span></div>
  <div class="main">
  {for:$i,0,$menu_u}
    <div class="cat1"><a href="{$menu[$i]['url']}">{$menu[$i]['word']}</a></div>
  {/for}
  </div>
</div>
<!-- 新秀 -->
