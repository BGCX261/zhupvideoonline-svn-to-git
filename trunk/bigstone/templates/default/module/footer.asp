<%response.end%>
<div id="footer">
  <div class="nav">
  {for:$i,0,$footer_nav_u}
    <a href="{$footer_nav[$i]['url']}">{$footer_nav[$i]['word']}</a>
    {if:$i,neq,$footer_nav_u}|{/if}
  {/for}
  </div>
  <div class="info">
  Powered by <a href="http://www.sinsiu.com/" target="_blank">sinsiu</a>
  <a href="{$sit_code_url}" target="_blank">{$sit_code}</a>
  <a href="{$sit_tec_url}" target="_blank">{$sit_tec}</a>
  {$sit_count_code}
  </div>
</div>
<!-- 新秀 -->
