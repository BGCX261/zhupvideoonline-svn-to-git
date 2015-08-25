<%response.end%>
<div class="box list" id="news">
  <div class="head">
    <div class="l"></div>
    <span>{$L_NEWS}</span>
    <a class="more" href="{$S_ROOT}?article.{$suffix}">{$L_MORE}</a>
    <div class="r"></div>
  </div>
  <ul class="main">
    {for:$i,0,$news_u}
    <li><a href="{$S_ROOT}?article_{$news[$i]['art_id']}.{$suffix}" target="_blank">{$cut_str($news[$i]['art_title'],14)}</a></li>
    {/for}
  </ul>
</div>
<!-- 新秀 -->
