<%response.end%>
<div class="here">
  <span>{$channel_title}</span>
  <div class="link">
  <a href="{$S_ROOT}">{$L_HOME}</a>
  {if:$channel_title,neq,""}<img src="{$S_TPL_PATH}images/here_img.jpg" /><a href="{$S_ROOT}?{$url_page}.{$suffix}">{$channel_title}</a>{/if}
  {if:$cat_name,neq,""}<img src="{$S_TPL_PATH}images/here_img.jpg" /><a href="{$S_ROOT}?{$url_page}_{$cat_id}_0_0.{$suffix}">{$cat_name}</a>{/if}
  {if:$page_title,neq,""}<img src="{$S_TPL_PATH}images/here_img.jpg" /><a href="{$S_ROOT}?{$url}">{$cut_str($page_title,20)}</a>{/if}
  </div>
</div>
<!-- 新秀 -->