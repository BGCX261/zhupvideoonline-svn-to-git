<%response.end%>
{if:$show_sheet,eq,1}
<div class="art_sheet list">
  <ul class="main">
    {for:$i,0,$article_u}
    <li><a href="{$S_ROOT}?article_{$article[$i]['art_id']}.{$suffix}" target="_blank">{$left($article[$i]['art_title'],50)}</a><span>{$datevalue($article[$i]['art_add_time'])}</span><div class="clear"></div></li>
    {/for}
  </ul>
  {if:$article_u,lt,0}<div class="not_found">{$L_NOT_FOUND}</div>{/if}
</div>
<div class="page_link"><div class="in">
  <span>{$L_ALL}{$page_sum}{$L_PAGES}</span>
  <span>{$L_NO}{$page}{$L_PAGE}</span>
  {if:$page_sum,neq,1}
    <a href="{$S_ROOT}?article_{$cat}_1_0.{$suffix}">{$L_FIRST_PAGE}</a>
    {if:$page-1,gt,0}<a href="{$S_ROOT}?article_{$cat}_{$page-1}_0.{$suffix}">{$L_PREVIOUS_PAGE}</a>{/if}
    {if:$page-4,gt,0}<a class="num" href="{$S_ROOT}?article_{$cat}_{$page-4}_0.{$suffix}">【{$page-4}】</a>{/if}
    {if:$page-3,gt,0}<a class="num" href="{$S_ROOT}?article_{$cat}_{$page-3}_0.{$suffix}">【{$page-3}】</a>{/if}
    {if:$page-2,gt,0}<a class="num" href="{$S_ROOT}?article_{$cat}_{$page-2}_0.{$suffix}">【{$page-2}】</a>{/if}
    {if:$page-1,gt,0}<a class="num" href="{$S_ROOT}?article_{$cat}_{$page-1}_0.{$suffix}">【{$page-1}】</a>{/if}
    <a class="num" href="{$S_ROOT}?article_{$cat}_{$page}_0.{$suffix}" style="color:#7fb402;">【{$page}】</a>
    {if:$page+1,elt,$page_sum}<a class="num" href="{$S_ROOT}?article_{$cat}_{$page+1}_0.{$suffix}">【{$page+1}】</a>{/if}
    {if:$page+2,elt,$page_sum}<a class="num" href="{$S_ROOT}?article_{$cat}_{$page+2}_0.{$suffix}">【{$page+2}】</a>{/if}
    {if:$page+3,elt,$page_sum}<a class="num" href="{$S_ROOT}?article_{$cat}_{$page+3}_0.{$suffix}">【{$page+3}】</a>{/if}
    {if:$page+4,elt,$page_sum}<a class="num" href="{$S_ROOT}?article_{$cat}_{$page+4}_0.{$suffix}">【{$page+4}】</a>{/if}
    {if:$page+1,elt,$page_sum}<a href="{$S_ROOT}?article_{$cat}_{$page+1}_0.{$suffix}">{$L_NEXT_PAGE}</a>{/if}
    <a href="{$S_ROOT}?article_{$cat}_{$page_sum}_0.{$suffix}">{$L_LAST_PAGE}</a>
  {/if}
  <form name="form_jump" action="" method="get">
    <input class="text" type="text" name="page" value="" />
    <input class="button" type="button" onclick="page_jump()" value="{$L_JUMP}"/>
  </form>
  <script language="javascript">
  function page_jump()
  {
	  var val = document.form_jump.page.value;
	  document.location.href = "{$S_ROOT}?article_{$cat}_" + val + "_0.{$suffix}";
  }
  </script>
</div></div>
{else}
<div id="article">
  <div class="title">
    <h3>{$article[0]['art_title']}</h3>
  </div>
  <div class="message">
    {$L_AUTHOR}：{$article[0]['att_author']}&nbsp;&nbsp;&nbsp;{$L_ADD_TIME}：{$article[0]['art_add_time']}
  </div>
  <div class="main">
    {$article[0]['art_text']}
  </div>
  <div id="prev_next">
    <div class="prev">{$L_PREV}：{if:$prev_title,neq,""}<a href="{$S_ROOT}?article_{$prev_id}.html">{$left($prev_title,20)}</a>{else}{$L_NONE}{/if}</div>
    <div class="next">{$L_NEXT}：{if:$next_title,neq,""}<a href="{$S_ROOT}?article_{$next_id}.html">{$left($next_title,20)}</a>{else}{$L_NONE}{/if}</div>
  </div>
</div>
{/if}
<!-- 新秀 -->
