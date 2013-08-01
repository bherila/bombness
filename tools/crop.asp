<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
 <p><strong>Crop Image</strong></p>
 <p>Select top-left dimension:</p>
 <form name="form1" method="get" action="crop-2.asp">
   <input name="a" type="image" id="a" src="/imgsrv?f=<%= Request("fname") %>" border="0">
   <input name="fname" type="hidden" id="fname" value="<%= Request("fname") %>"> 
</form>
 <p>&nbsp; </p>
</body>
</html>
