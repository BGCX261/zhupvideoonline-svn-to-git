<%response.end%>
<div class="mes_sheet">
  <ul>
    {for:$i,0,$message_u}
    <li>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="30"><span>{$message[$i]['mes_user']}</span>&nbsp;{$message[$i]['mes_time']}&nbsp;{$L_SAID}：</td>
        </tr>
        <tr>
          <td><div>{$message[$i]['mes_text']}</div></td>
        </tr>
        {if:$message[$i]['mes_reply'],neq,""}
        <tr>
          <td><div class="reply"><span>{$L_REPLY}：</span>{$message[$i]['mes_reply']}</div></td>
        </tr>
        {/if}
      </table>
    </li>
    {/for}
  </ul>
  {if:$message_u,lt,0}<div class="not_found">{$L_NOT_FOUND}</div>{/if}
</div>
<div class="page_link"><div class="in">
  <span>{$L_ALL}{$page_sum}{$L_PAGES}</span>
  <span>{$L_NO}{$page}{$L_PAGE}</span>
  {if:$page_sum,neq,1}
    <a href="{$S_ROOT}?message_{$cat}_1_0.{$suffix}">{$L_FIRST_PAGE}</a>
    {if:$page-1,gt,0}<a href="{$S_ROOT}?message_{$cat}_{$page-1}_0.{$suffix}">{$L_PREVIOUS_PAGE}</a>{/if}
    {if:$page-4,gt,0}<a class="num" href="{$S_ROOT}?message_{$cat}_{$page-4}_0.{$suffix}">【{$page-4}】</a>{/if}
    {if:$page-3,gt,0}<a class="num" href="{$S_ROOT}?message_{$cat}_{$page-3}_0.{$suffix}">【{$page-3}】</a>{/if}
    {if:$page-2,gt,0}<a class="num" href="{$S_ROOT}?message_{$cat}_{$page-2}_0.{$suffix}">【{$page-2}】</a>{/if}
    {if:$page-1,gt,0}<a class="num" href="{$S_ROOT}?message_{$cat}_{$page-1}_0.{$suffix}">【{$page-1}】</a>{/if}
    <a class="num" href="{$S_ROOT}?message_{$cat}_{$page}_0.{$suffix}" style="color:#2874c2;">【{$page}】</a>
    {if:$page+1,elt,$page_sum}<a class="num" href="{$S_ROOT}?message_{$cat}_{$page+1}_0.{$suffix}">【{$page+1}】</a>{/if}
    {if:$page+2,elt,$page_sum}<a class="num" href="{$S_ROOT}?message_{$cat}_{$page+2}_0.{$suffix}">【{$page+2}】</a>{/if}
    {if:$page+3,elt,$page_sum}<a class="num" href="{$S_ROOT}?message_{$cat}_{$page+3}_0.{$suffix}">【{$page+3}】</a>{/if}
    {if:$page+4,elt,$page_sum}<a class="num" href="{$S_ROOT}?message_{$cat}_{$page+4}_0.{$suffix}">【{$page+4}】</a>{/if}
    {if:$page+1,elt,$page_sum}<a href="{$S_ROOT}?message_{$cat}_{$page+1}_0.{$suffix}">{$L_NEXT_PAGE}</a>{/if}
    <a href="{$S_ROOT}?message_{$cat}_{$page_sum}_0.{$suffix}">{$L_LAST_PAGE}</a>
  {/if}
  <form name="form_jump" action="" method="get">
    <input class="text" type="text" name="page" value="" />
    <input class="button" type="button" onclick="page_jump()" value="{$L_JUMP}"/>
  </form>
  <script language="javascript">
  function page_jump()
  {
	  var val = document.form_jump.page.value;
	  document.location.href = "{$S_ROOT}?message_{$cat}_" + val + "_0.{$suffix}";
  }
  </script>
</div></div>
<div id="leave_word">
  <form action="{$S_ROOT}index/submit.asp" method="post">
    <input name="cmd" type="hidden" value="leave_word"/>
    <div>
      <input name="user" class="text" type="text" maxlength="30" value="{$L_TRAVELLER}"/>
      <input name="show" type="checkbox" value="2"/>{$L_SECRET}
    </div>
    <div>
      <textarea name="text" cols="60" rows="5"></textarea>
    </div>
    <div>
      <input class="button" name="submit" type="submit" value="{$L_SUBMIT}" />
      <input class="button" name="reset" type="reset" value="{$L_RESET}"/>
    </div>
  </form>
</div>