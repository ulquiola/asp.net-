<html>
<head>
<title>选择模板</title>
<script language="JavaScript">
<!--
	function return_type( t, s ) {

		if (!opener) return true;
		with( opener.document.SelectTempletForm ) {
			switch( t.toLowerCase() ) {
				case "a" :
					for(var i=0; i<TempletNum_A.options.length; i++) {
						if( TempletNum_A.options[i].value == s )
							TempletNum_A.options[i].selected = true;
					}
			}
			self.close();
			return false;
		}
	}
//-->
</script>
<style type="text/css">
<!--
.pt105 {  font-size: 10.5pt}
-->
</style>
</head>
<body bgcolor="#F1F1F1" topmargin=10 marginheight=0 leftmargin=0 marginwidth=0 width="100%">
<form name="" action="" method="post">
  <table border=0 cellspacing=3 cellpadding=3 align=center>
    <tr>
      <td><table border=0 cellspacing=3 cellpadding=3 align=center>
          <tr>
            <td align=center valign=middle width=152><table width="20" border="0" cellspacing="1" cellpadding="0" bgcolor="#333333">
                <tr>
                  <td bgcolor="#F1F1F1"><a href="image/1.gif" target="_blank"><img src="image/1.gif" boreer=0 class="BBImage" width=220 height=176 border="0" alt="点击放大"></a></td>
                </tr>
              </table></td>
            <td colspan="2" align=center valign=middle><table width="20" border="0" cellspacing="1" cellpadding="0" bgcolor="#333333">
                <tr>
                  <td bgcolor="#F1F1F1"><a href="image/2.gif" target="_blank"><img src="image/2.gif" boreer=0 class="BBImage" width=217 height=176 border="0" alt="点击放大"></a></td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td align=center><a href="javascript:return_type( 'a', 1 );"> <font face=verdana><b>1. </b></font>选择</a></td>
            <td width="152" align=center><a href="javascript:return_type( 'a', 2 );"> <font face=verdana><b>2. </b></font>选择</a></td>
          </tr>
          <tr>
            <td height="80" valign="middle" colspan="3"><div align="right" class="pt105"></div></td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>
