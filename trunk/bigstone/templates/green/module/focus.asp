<%response.end%>
<div id="focus_hull">
  <div id="focus">
    {if:$i,egt,0}
    <div id="focus_bg"></div>
    <div id="focus_show"></div>
    <div id="focus_img">
    {for:$i,0,$focus_u}
      <div name="focus_img" id="focus_{$i+1}">{$focus[$i]['path']}|{$focus[$i]['url']}|{$focus[$i]['title']}</div>
    {/for}
    </div>
    <script type="text/javascript" src="{$S_TPL_PATH}js/focus.js"></script>
    {/if}
  </div>
</div>
<!-- 新秀 -->
