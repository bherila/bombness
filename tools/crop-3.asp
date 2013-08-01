<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
 <p><strong>Crop Image</strong></p>
 <p>Is this ok? </p>
 <form name="form1" method="get" action="crop-3.aspx">
   <p>
<img src="crop-3.aspx?fname=<%= Request("fname") %>&a.x=<%= Request("a.x") %>&a.y=<%= Request("a.y") %>&b.x=<%= Request("b.x") %>&b.y=<%= Request("b.y")%>&prev=1"><input name="a.x" type="hidden" id="a.x" value="<%= Request("a.x") %>">
     <input name="a.y" type="hidden" id="a.y2" value="<%= Request("a.y") %>">
     <input name="b.y" type="hidden" id="a.y3" value="<%= Request("b.y") %>">
     <input name="b.x" type="hidden" id="a.y4" value="<%= Request("b.x") %>"> 
     <input name="fname" type="hidden" id="fname2" value="<%= Request("fname") %>">
     <input name="prev" type="hidden" id="fname3" value="0">
</p>
   <p>
     <input type="submit" name="Submit" value="Yes">
</p>
 </form>
 <p>&nbsp; </p>
</body>
</html>
