<%@ Language=VBScript codepage=65001 %>
<!-- #include file="secure.inc" -->
<%
set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")
set r2=Server.CreateObject("ADODB.Recordset")
	indexField = ""
	subTable = ""

thetable = request("page")
select case thetable
case "[stamps].[dbo].[valueData]"
	numOfFields = 2
	currpage = "[stamps].[dbo].[MainTable]"
	indexField = "theIndex"
	fieldWhere = "mainIndex"
	field1 = "coinType"
    field2 = "value"
	ftype1 = "select"
    ftype2 = "number"
	fname1 = "סוג מטבע"
    fname2 = "ערך"
case "[stamps].[dbo].[designData]"
	numOfFields = 2
	currpage = "[stamps].[dbo].[MainTable]"
	indexField = "theIndex"
	fieldWhere = "mainIndex"
	field1 = "designer"
    field2 = "painter"
	ftype1 = "select"
    ftype2 = "select"
	fname1 = "מעצב"
    fname2 = "צייר"
case "[stamps].[dbo].[category]"
	numOfFields = 1
	currpage = "[stamps].[dbo].[MainTable]"
	indexField = "theIndex"
	fieldWhere = "mainIndex"
	field1 = "cat"
    ftype1 = "select"
    fname1 = "קטגוריה"
case "[stamps].[dbo].[productionData]"
	numOfFields = 8
	currpage = "[stamps].[dbo].[MainTable]"
	indexField = "theIndex"
	fieldWhere = "mainIndex"
	field1 = "printMethod"
    field2 = "color"
    field3 = "phos"
    field4 = "waterMark"
    field5 = "perforation"
    field6 = "engravingNumber"
    field7 = "itemHeight"
    field8 = "itemWidth"
	ftype1 = "select"
    ftype2 = "select"
    ftype3 = "select"
    ftype4 = "select"
    ftype5 = "select"
    ftype6 = "number"
    ftype7 = "number"
    ftype8 = "number"
	fname1 = "שיטת הדפסה"
    fname2 = "צבע"
    fname3 = "זרחן"
    fname4 = "סימן מים"
    fname5 = "ניקוב"
    fname6 = "מספר גלופה"
    fname7 = "גובה"
    fname8 = "רוחב"
end select

if request("sumbitButton")<>"" then
'sql2 = "select * from pages where num="&request("num")
'r2.open sql2,strconn,1,3
'if not r2.eof then theFolder=r2("page_title")
'r2.close

fields = field1
fields_content = "'"&replace(request(field1),"'","&#8217;")&"'"
for i = 2 to numOfFields
	if eval("field"&i)<>"" then
	curr_fields = replace(request(eval("field"&i)),"'","&#8217;")
		if (eval("ftype"&i)="number" or eval("ftype"&i)="date") and curr_fields="" then
			fields = fields
		else
			fields = fields&" ,["&eval("field"&i)&"]"
			if eval("ftype"&i) = "checkbox" then
				if curr_fields = "on" then
					curr_fields = true
				else
					curr_fields = false
				end if
			else
			curr_fields = "'"&curr_fields&"'"
			end if
			fields_content = fields_content&" ,"&curr_fields
		end if
	end if
next
	sql = "INSERT INTO "&thetable&" (["&fieldWhere&"], "&fields&") VALUES ('"&request("num")&"', "&fields_content&")"
'response.Write sql
'response.End
r.open sql,strconn,1,3
'pos = InStr(thetable,"_")
'page = left(thetable,pos-1)
%>
<script>
location.href= "edit.asp?page=<%=currpage%>&num=<%=request("num")%>&all=3"
</script>
<%
end if


sql = "SELECT * FROM "&thetable
r.open sql,strconn,1,3

function selectFunction(theField)
	select case theField
	case "coinType":
		sql1 = "SELECT * FROM [stamps].[dbo].[coinType]"
		r1.open sql1,strconn,1,3
		response.Write "<select name="&theField&"><option value='0'>בחר...</option>"
		while not r1.eof
		response.Write "<option value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		r1.MoveNext
		wend
		response.Write "</select>"
        r1.close
    case "designer":
		sql1 = "SELECT * FROM [stamps].[dbo].[designer]"
		r1.open sql1,strconn,1,3
		response.Write "<select name="&theField&"><option value='0'>בחר...</option>"
		while not r1.eof
		response.Write "<option value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		r1.MoveNext
		wend
		response.Write "</select>"
        r1.close
    case "painter":
		sql1 = "SELECT * FROM [stamps].[dbo].[designer]"
		r1.open sql1,strconn,1,3
		response.Write "<select name="&theField&"><option value='0'>בחר...</option>"
		while not r1.eof
		response.Write "<option value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		r1.MoveNext
		wend
		response.Write "</select>"
        r1.close
    case "printMethod":
		sql1 = "SELECT * FROM [stamps].[dbo].[printMethod]"
		r1.open sql1,strconn,1,3
		response.Write "<select name="&theField&"><option value='0'>בחר...</option>"
		while not r1.eof
		response.Write "<option value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		r1.MoveNext
		wend
		response.Write "</select>"
        r1.close
    case "color":
		sql1 = "SELECT * FROM [stamps].[dbo].[stampColor]"
		r1.open sql1,strconn,1,3
		response.Write "<select name="&theField&"><option value='0'>בחר...</option>"
		while not r1.eof
		response.Write "<option value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		r1.MoveNext
		wend
		response.Write "</select>"
        r1.close
    case "phos":
		sql1 = "SELECT * FROM [stamps].[dbo].[phosphor]"
		r1.open sql1,strconn,1,3
		response.Write "<select name="&theField&"><option value='0'>בחר...</option>"
		while not r1.eof
		response.Write "<option value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		r1.MoveNext
		wend
		response.Write "</select>"
        r1.close
    case "waterMark":
		sql1 = "SELECT * FROM [stamps].[dbo].[waterMark]"
		r1.open sql1,strconn,1,3
		response.Write "<select name="&theField&"><option value='0'>בחר...</option>"
		while not r1.eof
		response.Write "<option value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		r1.MoveNext
		wend
		response.Write "</select>"
        r1.close
    case "perforation":
		sql1 = "SELECT * FROM [stamps].[dbo].[perforation]"
		r1.open sql1,strconn,1,3
		response.Write "<select name="&theField&"><option value='0'>בחר...</option>"
		while not r1.eof
		response.Write "<option value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		r1.MoveNext
		wend
		response.Write "</select>"
        r1.close
    case "cat":
		sql1 = "SELECT * FROM [stamps].[dbo].[stampCategories]"
		r1.open sql1,strconn,1,3
		response.Write "<select name="&theField&"><option value='0'>בחר...</option>"
		while not r1.eof
		response.Write "<option value="&r1("RecID")&">"&r1("CategoryName_HE")&"</option>"
		r1.MoveNext
		wend
		response.Write "</select>"
        r1.close
	end select
		
end function
%>

<!---------------------------------------------------------------------------------------------------------------------------------
 #################################################################################################################################
 ###
 ###   TODO ON THIS FILE:
 ###   inside the select case - 
 ###
 ###   case "TABLE NAME"
 ###   numOfFields = NUMBER OF FIELDS TO EDIT IN THE TABLE
 ###   currpage = PARENT TABLE
 ###   indexField = "NAME OF KEY FIELD OF THE TABLE"
 ###   fieldWhere = "THE FIELD THAT CONACTED TO THE PARENT TABLE"
 ###   field1, 2, 3... = "THE NAME OF THE FIELD IN THE TABLE"
 ###   ftype1, 2, 3... = "THE FIELD'S TYPE"
 ###   fname1, 2, 3... = "TEXT TO PRESENT FOR THE FIELD"
 ###
 ###   RULES:
 ###   NAMES OF TABLES AND FIELDS CAN'T BE: page, table, where, field, text, date, number, on ,off, name, add, edit
 ###   TYPES OF FIELDS: text; number; date; checkbox; *select; **browse;
 ###   * select field calls the selectFunction and should be adjusted per use if necessary
 ###   ** browse field calls the upload.asp and upload1.asp files, should be adjusted per server
 ###
 #################################################################################################################################
------------------------------------------------------------------------------------------------------------------------------------>

<html><head><title>מערכת ניהול</title>
<!-- #INCLUDE file="FCKeditor/fckeditor.asp" -->
<script language=javascript>
function checkForm(theForm)
{
	if(theForm.<%=field1%>.value=="")
	{
		document.getElementById("errorTr").style.display="inline";
		return false;
	}
	else
	{
		document.getElementById("errorTr").style.display="none";
		return true;
	}
}
</script>
<meta charset="UTF-8">
</head>
<body LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
<link rel="stylesheet" type="text/css" href="dhtmlgoodies_calendar.css">
<script language=javascript src="dhtmlgoodies_calendar.js"></script>
<link rel="stylesheet" type="text/css" href="admin.css">
<table class=text_blue cellspacing=0 cellpadding=0 align=center ID="Table1"><tr><td dir=rtl>
</td></tr>
<tr align=right><td height=50 class=text_white colspan=2 bgcolor="#990000" align=center>מערכת ניהול</td></tr>
<tr><td width=1 bgcolor="#990000"></td><td>
<table class=text_blue cellspacing=0 cellpadding=0 ID="Table3"><tr><td align=right valign=top width=616 height=200>
<table class=text_blue ID="Table5" cellpadding=5 cellspacing=0 border=0><tr><td width=20>&nbsp;</td>
<td valign=top align=right dir=rtl>
<form method=post onsubmit="return checkForm(this)">
<table cellpadding=0 cellspacing=0 class=text_blue ID="Table4" border=0 dir=rtl>
<tr style="display:none" id="errorTr"><td colspan=2 class=oblig>נא למלא <%=fname1%></td></tr>
<%for i=1 to numOfFields
if i = 1 then
	oblig = "<span class='oblig'>*</span>"
else
	oblig = ""
end if
	%>
	<tr><td style="padding:10px" valign=top><%=oblig%><%=eval("fname"&i)%></td><td dir=rtl valign=top style="padding:5">
	<%if eval("ftype"&i) = "date" then%>
		<nobr><input readonly onclick="displayCalendar(this,'dd/mm/yyyy',this)" type=text ID="Text3" NAME="<%=eval("field"&i)%>"> <input class="submit" type="button" style="width:70" value="תאריכון" onclick="displayCalendar(document.forms[0].<%=eval("field"&i)%>,'dd/mm/yyyy',this)" ID="Button1" NAME="Button1"></nobr>
	<%end if%>
	<%if eval("ftype"&i) = "text" then%>
		<input type=text ID="Text1" NAME="<%=eval("field"&i)%>">
	<%end if%>
	<%if eval("ftype"&i) = "browse" then
        sql2 = "select * from pages where num="&request("num")
        r2.open sql2,strconn,1,3
        if not r2.eof then theFolder=r2("page_title")
        r2.close
        %>
		<input type=text NAME="<%=eval("field"&i)%>" readonly ID="File1">&nbsp;&nbsp;<input value="דפדף..." type=button onclick="window.open('upload.asp?fieldname=<%=eval("field"&i)%>&thefolder=<%=theFolder%>','','width=270, height=70')" ID="Button2" NAME="Button2">
	<%end if%>
	<%if eval("ftype"&i) = "number" then%>
		<input type=text onkeypress="if(event.keyCode<48 || event.keyCode>57) return false" ID="Text4" NAME="<%=eval("field"&i)%>">
	<%end if%>
	<%if eval("ftype"&i) = "textarea" then%>
		<textarea rows=10 cols=35 ID=textarea NAME="<%=eval("field"&i)%>"></textarea>
	<%end if%>
	<%if eval("ftype"&i) = "checkbox" then%>
		<input type=checkbox ID="Text2" NAME="<%=eval("field"&i)%>">
	<%end if%> 
	<%if eval("ftype"&i) = "select" then
		selectFunction(eval("field"&i))
	end if%> 
	</td></tr>
<%next
%>
<tr><td colspan=2 align=center><input class="submit" style="width:50px" type="submit" value="הוסף" name=sumbitButton></td></tr>
</table>
</form>
</td></tr></table>
</td><td width=1 bgcolor="#990000"></td></tr><tr align=right><td height=1 class=text_white colspan=10 bgcolor=#990000 align=center></td></tr></table>
</td></tr></table>