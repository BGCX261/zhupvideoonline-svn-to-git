<%response.end%>
{if:$show_sheet,eq,1}
<div class="img_sheet">
  {for:$i,0,$product_u}
  <div class="unit"><div class="in">
    <table cellspacing="0" cellpadding="0">
      <tr><td class="img">
        <a href="{$S_ROOT}?product_{$product[$i]['art_id']}.{$suffix}" target="_blank"><img src="{$S_ROOT}{$product[$i]['art_reduce_img']}" onload="picresize(this,150,150)"/></a>
      </td></tr>
      <tr><td class="title">
        <a href="{$S_ROOT}?product_{$product[$i]['art_id']}.{$suffix}" target="_blank">{$cut_str($product[$i]['art_title'],11)}</a>
      </td></tr>
    </table>
  </div></div>
  {/for}
  {if:$product_u,lt,0}<div class="not_found">{$L_NOT_FOUND}</div>{/if}
</div>
<div class="page_link"><div class="in">
  <span>{$L_ALL}{$page_sum}{$L_PAGES}</span>
  <span>{$L_NO}{$page}{$L_PAGE}</span>
  {if:$page_sum,neq,1}
    <a href="{$S_ROOT}?product_{$cat}_1_0.{$suffix}">{$L_FIRST_PAGE}</a>
    {if:$page-1,gt,0}<a href="{$S_ROOT}?product_{$cat}_{$page-1}_0.{$suffix}">{$L_PREVIOUS_PAGE}</a>{/if}
    {if:$page-4,gt,0}<a class="num" href="{$S_ROOT}?product_{$cat}_{$page-4}_0.{$suffix}">【{$page-4}】</a>{/if}
    {if:$page-3,gt,0}<a class="num" href="{$S_ROOT}?product_{$cat}_{$page-3}_0.{$suffix}">【{$page-3}】</a>{/if}
    {if:$page-2,gt,0}<a class="num" href="{$S_ROOT}?product_{$cat}_{$page-2}_0.{$suffix}">【{$page-2}】</a>{/if}
    {if:$page-1,gt,0}<a class="num" href="{$S_ROOT}?product_{$cat}_{$page-1}_0.{$suffix}">【{$page-1}】</a>{/if}
    <a class="num" href="{$S_ROOT}?product_{$cat}_{$page}_0.{$suffix}" style="color:#d90000;">【{$page}】</a>
    {if:$page+1,elt,$page_sum}<a class="num" href="{$S_ROOT}?product_{$cat}_{$page+1}_0.{$suffix}">【{$page+1}】</a>{/if}
    {if:$page+2,elt,$page_sum}<a class="num" href="{$S_ROOT}?product_{$cat}_{$page+2}_0.{$suffix}">【{$page+2}】</a>{/if}
    {if:$page+3,elt,$page_sum}<a class="num" href="{$S_ROOT}?product_{$cat}_{$page+3}_0.{$suffix}">【{$page+3}】</a>{/if}
    {if:$page+4,elt,$page_sum}<a class="num" href="{$S_ROOT}?product_{$cat}_{$page+4}_0.{$suffix}">【{$page+4}】</a>{/if}
    {if:$page+1,elt,$page_sum}<a href="{$S_ROOT}?product_{$cat}_{$page+1}_0.{$suffix}">{$L_NEXT_PAGE}</a>{/if}
    <a href="{$S_ROOT}?product_{$cat}_{$page_sum}_0.{$suffix}">{$L_LAST_PAGE}</a>
  {/if}
  <form name="form_jump" action="" method="get">
    <input class="text" type="text" name="page" value="" />
    <input class="button" type="button" onclick="page_jump()" value="{$L_JUMP}"/>
  </form>
  <script language="javascript">
  function page_jump()
  {
	  var val = document.form_jump.page.value;
	  document.location.href = "{$S_ROOT}?product_{$cat}_" + val + "_0.{$suffix}";
  }
  </script>
</div></div>
{else}
<div id="picture">
  <div class="img">
    <table cellspacing="0" cellpadding="0">
      <tr><td>
        <img src="{$S_ROOT}{$product[0]['art_img']}"/>
      </td></tr>
    </table>
  </div>
  <div class="head">{$L_PRODUCT_ATTRIBUTE}</div>
  <div class="attribute">
    <table cellspacing="0" cellpadding="3">
      <tr>
        <td width="70"><span>{$L_PRODUCT_NAME}：</span></td>
        <td>{$product[0]['art_title']}</td>
      </tr>
      <tr>
        <td><span>{$L_PRODUCT_NUMBER}：</span></td>
        <td>{$product[0]['att_number']}</td>
      </tr>
      <tr>
        <td><span>{$L_PRODUCT_PRICE}：</span></td>
        <td>{$product[0]['att_price']}</td>
      </tr>
    </table>
  </div>
  <div class="head">{$L_PRODUCT_TEXT}</div>
  <div class="text">{$product[0]['art_text']}</div>
  <div id="prev_next">
    <div class="prev">{$L_PREV}：{if:$prev_title,neq,""}<a href="{$S_ROOT}?product_{$prev_id}.html">{$cut_str($prev_title,20)}</a>{else}{$L_NONE}{/if}</div>
    <div class="next">{$L_NEXT}：{if:$next_title,neq,""}<a href="{$S_ROOT}?product_{$next_id}.html">{$cut_str($next_title,20)}</a>{else}{$L_NONE}{/if}</div>
    <div class="clear"></div>
  </div>
</div>
{/if}
<!-- 新秀 -->
