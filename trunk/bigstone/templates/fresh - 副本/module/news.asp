<%response.end%>
<div class="box2 list" id="news">
  <div class="head"><span>{$L_NEWS}</span><a class="more" href="{$S_ROOT}?article.{$suffix}"><img src="{$S_TPL_PATH}images/more.gif" /></a></div>
  <ul class="main">
    {for:$i,0,$news_u}
    <li><a href="{$S_ROOT}?article_{$news[$i]['art_id']}.{$suffix}" target="_blank">{$left($news[$i]['art_title'],20)}</a><span>{$datevalue($news[$i]['art_add_time'])}</span><div class="clear"></div></li>
    {/for}
  </ul>
</div>
<!-- 新秀 -->
