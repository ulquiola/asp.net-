<div id="header"><img src="images/banner.gif"></div>
<div id="menu">
  <form id="form1" name="form1" method="post" action="" style="margin:1px;">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <a href="index.asp"><img src="images/btn_index.gif" width="87" height="22" border="0" /></a>
	<a href="#" onclick="loadScripAdd_window()"><img src="images/btn_music.gif" alt="�����б�" width="87" height="22" border="0" /></a>
	<a href="message_list.asp"><img src="images/btn_list.gif" alt="�����б�" border="0" /></a>
	<a href="message_list.asp?flag=lei"><img src="images/rank.gif" alt="������" border="0" /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;	
	<img src="images/text_search.gif" alt="����" />
    <input name="keyID" id="keyID" class="input" type="text" value="�������" size="7" onclick="this.value = '';" />
	<input type="image" src="images/btn_search.gif" onclick="searchScrip(this.form.keyID.value);return false;" />
 <script language="javascript">
function searchScrip(n){
	if(document.getElementById("scrip"+n)){
		Show(n,'shadeDiv');
	}else{
		if(form1.keyID.value=="")
		{alert("������������ţ�");form1.keyID.focus();return false;}
		alert("�����������������!");
	}
}
</script>
</form>  
</div>


