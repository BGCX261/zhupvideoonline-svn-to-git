<%response.end%>
<div class="rec_sheet">
  <ul>
    {for:$i,0,$recruit_u}
    <li>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td class="title"><div>{$recruit[$i]['art_title']}</div></td>
          <td class="time">{$recruit[$i]['art_add_time']}</td>
        </tr>
        <tr>
          <td colspan="2" class="text">{$recruit[$i]['art_text']}</td>
        </tr>
      </table>
    </li>
    {/for}
  </ul>
  {if:$recruit_u,lt,0}<div class="not_found">{$L_NOT_FOUND}</div>{/if}
</div>
<div class="page_link"><div class="in">
  <span>{$L_ALL}{$page_sum}{$L_PAGES}</span>
  <span>{$L_NO}{$page}{$L_PAGE}</span>
  {if:$page_sum,neq,1}
    <a href="{$S_ROOT}?recruit_{$cat}_1_0.{$suffix}">{$L_FIRST_PAGE}</a>
    {if:$page-1,gt,0}<a href="{$S_ROOT}?recruit_{$cat}_{$page-1}_0.{$suffix}">{$L_PREVIOUS_PAGE}</a>{/if}
    {if:$page-4,gt,0}<a class="num" href="{$S_ROOT}?recruit_{$cat}_{$page-4}_0.{$suffix}">【{$page-4}】</a>{/if}
    {if:$page-3,gt,0}<a class="num" href="{$S_ROOT}?recruit_{$cat}_{$page-3}_0.{$suffix}">【{$page-3}】</a>{/if}
    {if:$page-2,gt,0}<a class="num" href="{$S_ROOT}?recruit_{$cat}_{$page-2}_0.{$suffix}">【{$page-2}】</a>{/if}
    {if:$page-1,gt,0}<a class="num" href="{$S_ROOT}?recruit_{$cat}_{$page-1}_0.{$suffix}">【{$page-1}】</a>{/if}
    <a class="num" href="{$S_ROOT}?recruit_{$cat}_{$page}_0.{$suffix}" style="color:#7fb402;">【{$page}】</a>
    {if:$page+1,elt,$page_sum}<a class="num" href="{$S_ROOT}?recruit_{$cat}_{$page+1}_0.{$suffix}">【{$page+1}】</a>{/if}
    {if:$page+2,elt,$page_sum}<a class="num" href="{$S_ROOT}?recruit_{$cat}_{$page+2}_0.{$suffix}">【{$page+2}】</a>{/if}
    {if:$page+3,elt,$page_sum}<a class="num" href="{$S_ROOT}?recruit_{$cat}_{$page+3}_0.{$suffix}">【{$page+3}】</a>{/if}
    {if:$page+4,elt,$page_sum}<a class="num" href="{$S_ROOT}?recruit_{$cat}_{$page+4}_0.{$suffix}">【{$page+4}】</a>{/if}
    {if:$page+1,elt,$page_sum}<a href="{$S_ROOT}?recruit_{$cat}_{$page+1}_0.{$suffix}">{$L_NEXT_PAGE}</a>{/if}
    <a href="{$S_ROOT}?recruit_{$cat}_{$page_sum}_0.{$suffix}">{$L_LAST_PAGE}</a>
  {/if}
  <form name="form_jump" action="" method="get">
    <input class="text" type="text" name="page" style="width:30px;" value="" />
    <input class="button" type="button" onclick="page_jump()" value="{$L_JUMP}"/>
  </form>
  <script language="javascript">
  function page_jump()
  {
	  var val = document.form_jump.page.value;
	  document.location.href = "{$S_ROOT}?recruit_{$cat}_" + val + "_0.{$suffix}";
  }
  </script>
</div></div>
<!-- 新秀 -->
