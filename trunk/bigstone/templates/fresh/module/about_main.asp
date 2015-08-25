<%response.end%>
<div class="here">
  <span>{if:$page_title,neq,""}{$left($page_title,10)}{else}{$channel_title}{/if}</span>
  <div class="link">
  <a href="{$S_ROOT}">{$L_HOME}</a>
  {if:$channel_title,neq,""}<img src="{$S_TPL_PATH}images/here_img.gif" /><a href="{$S_ROOT}?{$url_page}.{$suffix}">{$channel_title}</a>{/if}
  {if:$page_title,neq,""}<img src="{$S_TPL_PATH}images/here_img.gif" /><a href="{$S_ROOT}?{$url}">{$left($page_title,20)}</a>{/if}
  </div>
</div>
<div id="about_main">
  {if:$about_u,egt,0}{$about[0]['art_text']}{/if}
</div>
<!-- 新秀 -->
