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
sql = "UPDATE "&thetable&" set "
for i=1 to numOfFields
	if eval("field"&i)<>"" then
		if i>1 then
			sql = sql&" ,"
		end if
		tablevalue = replace(request(eval("field"&i)),"'","&#8217;")
		if tablevalue = "" and eval("ftype"&i) = "number" then
			tablevalue = "0"
		end if
		if tablevalue = "" and eval("ftype"&i) = "text" then
			tablevalue = " "
		end if
		if tablevalue = "" and eval("ftype"&i) = "date" then
			tablevalue = "1/1/1"
		end if
		if eval("ftype"&i) = "checkbox" then
			if request(eval("field"&i)) = "on" then
				tablevalue = true
			else
				tablevalue = false
			end if
		else
		tablevalue = "'"&tablevalue&"'"
		end if
		
		sql = sql&"["&eval("field"&i)&"] = "&tablevalue
	end if
next
sql = sql &" WHERE "&indexField&" = "&request("num")
r.open sql,strconn,1,3
end if


if request("del") <> "" then
    theNumber=0
    sql = "select * FROM "&thetable&" WHERE "&indexField&" = "&request("del")
    r1.open sql,strconn,1,3
    if not r1.eof then theNumber=r1("mainIndex")
    r1.close
	sql = "DELETE FROM "&thetable&" WHERE "&indexField&" = "&request("del")
	r1.open sql,strconn,1,3
	'response.Write sql
	response.Redirect "edit.asp?page="&currpage&"&num="&theNumber
	response.End
end if


sql = "SELECT * FROM "&thetable&" WHERE "&indexField&" = "&request("num")&""
'response.write sql
r.open sql,strconn,1,3
prev_page = "page="&currpage&"&num="&r(fieldWhere)
prev_num = r(fieldWhere)

sub selectFunction(theField, fieldValue)
    select case theField
	case "coinType":
		sql1 = "SELECT * FROM [stamps].[dbo].[coinType]"
		r1.open sql1,strconn,1,3
		response.Write "<select name="&theField&"><option value='0'>בחר...</option>"
		while not r1.eof
		    if int(r1("oldIndex")) = int(fieldValue) then
			    response.Write "<option selected value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		    else
			    response.Write "<option value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		    end if
		r1.MoveNext
		wend
		response.Write "</select>"
        r1.close
    case "designer":
		sql1 = "SELECT * FROM [stamps].[dbo].[designer]"
		r1.open sql1,strconn,1,3
		response.Write "<select name="&theField&"><option value='0'>בחר...</option>"
		while not r1.eof
		    if int(r1("oldIndex")) = int(fieldValue) then
			    response.Write "<option selected value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		    else
			    response.Write "<option value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		    end if
		r1.MoveNext
		wend
		response.Write "</select>"
        r1.close
    case "painter":
		sql1 = "SELECT * FROM [stamps].[dbo].[designer]"
		r1.open sql1,strconn,1,3
		response.Write "<select name="&theField&"><option value='0'>בחר...</option>"
		while not r1.eof
		    if int(r1("oldIndex")) = int(fieldValue) then
			    response.Write "<option selected value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		    else
			    response.Write "<option value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		    end if
		r1.MoveNext
		wend
		response.Write "</select>"
        r1.close
    case "printMethod":
		sql1 = "SELECT * FROM [stamps].[dbo].[printMethod]"
		r1.open sql1,strconn,1,3
		response.Write "<select name="&theField&"><option value='0'>בחר...</option>"
		while not r1.eof
		    if int(r1("oldIndex")) = int(fieldValue) then
			    response.Write "<option selected value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		    else
			    response.Write "<option value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		    end if
		r1.MoveNext
		wend
		response.Write "</select>"
        r1.close
    case "color":
		sql1 = "SELECT * FROM [stamps].[dbo].[stampColor]"
		r1.open sql1,strconn,1,3
		response.Write "<select name="&theField&"><option value='0'>בחר...</option>"
		while not r1.eof
		    if int(r1("oldIndex")) = int(fieldValue) then
			    response.Write "<option selected value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		    else
			    response.Write "<option value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		    end if
		r1.MoveNext
		wend
		response.Write "</select>"
        r1.close
    case "phos":
		sql1 = "SELECT * FROM [stamps].[dbo].[phosphor]"
		r1.open sql1,strconn,1,3
		response.Write "<select name="&theField&"><option value='0'>בחר...</option>"
		while not r1.eof
		    if int(r1("oldIndex")) = int(fieldValue) then
			    response.Write "<option selected value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		    else
			    response.Write "<option value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		    end if
		r1.MoveNext
		wend
		response.Write "</select>"
        r1.close
    case "waterMark":
		sql1 = "SELECT * FROM [stamps].[dbo].[waterMark]"
		r1.open sql1,strconn,1,3
		response.Write "<select name="&theField&"><option value='0'>בחר...</option>"
		while not r1.eof
		    if int(r1("oldIndex")) = int(fieldValue) then
			    response.Write "<option selected value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		    else
			    response.Write "<option value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		    end if
		r1.MoveNext
		wend
		response.Write "</select>"
        r1.close
    case "perforation":
		sql1 = "SELECT * FROM [stamps].[dbo].[perforation]"
		r1.open sql1,strconn,1,3
		response.Write "<select name="&theField&"><option value='0'>בחר...</option>"
		while not r1.eof
		    if int(r1("oldIndex")) = int(fieldValue) then
			    response.Write "<option selected value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		    else
			    response.Write "<option value="&r1("oldIndex")&">"&r1("hName")&"</option>"
		    end if
		r1.MoveNext
		wend
		response.Write "</select>"
        r1.close
    case "cat":
		sql1 = "SELECT * FROM [stamps].[dbo].[stampCategories]"
		r1.open sql1,strconn,1,3
		response.Write "<select name="&theField&"><option value='0'>בחר...</option>"
		while not r1.eof
		    if int(r1("RecID")) = int(fieldValue) then
			    response.Write "<option selected value="&r1("RecID")&">"&r1("CategoryName_HE")&"</option>"
		    else
			    response.Write "<option value="&r1("RecID")&">"&r1("CategoryName_HE")&"</option>"
		    end if
		r1.MoveNext
		wend
		response.Write "</select>"
        r1.close
	end select
end sub
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

function delete_sub()
{
	if(confirm("?למחוק רשומה"))
	    location.href="edit_sub.asp?del=<%=request("num")%>&page=<%=request("page")%>";
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
<table class=text_blue cellspacing=0 cellpadding=0 ID="Table3"><tr><td align=right valign=top background="../cut/pattern_g.jpg" width=616 height=200>
<table class=text_blue ID="Table5" cellpadding=5 cellspacing=0 border=0><tr><td width=15>&nbsp;</td>
<td valign=top align=right dir=rtl>
</td>
<td valign=top>
<form method=post onsubmit="return checkForm(this)">
<table cellpadding=0 cellspacing=0 class=text_blue ID="Table4" border=0>
<tr style="display:none" id="errorTr"><td colspan=2 class=oblig>נא למלא <%=fname1%></td></tr>
<%for i=1 to numOfFields
if i = 1 then
	oblig = "<span class='oblig'>*</span>"
else
	oblig = ""
end if
	%>
	<tr><td dir=rtl valign=top style="padding:10px; padding-top:0">
	<%if eval("ftype"&i) = "date" then%>
		<nobr><input readonly onclick="displayCalendar(this,'dd/mm/yyyy',this)" value="<%=r(eval("field"&i))%>" type=text ID="Text3" NAME="<%=eval("field"&i)%>"> <input type="button" class="submit" style="width:70" value="תאריכון" onclick="displayCalendar(document.forms[0].<%=eval("field"&i)%>,'dd/mm/yyyy',this)" ID="Button1" NAME="Button1"></nobr>
	<%end if%>
	<%if eval("ftype"&i) = "text" or eval("ftype"&i) = "number" then
		thevalue = r(eval("field"&i))
		if thevalue<>"" then
			thevalue = replace(thevalue,"""","&quot;")
		end if%>
		<input type=text <%if eval("ftype"&i) = "number" then%> onkeypress="if((event.keyCode<48 || event.keyCode>57) && event.keyCode!=46) return false" <%end if%> value="<%=thevalue%>" ID="Text1" NAME="<%=eval("field"&i)%>">
	<%end if%>
	<%if eval("ftype"&i) = "select" then
		call selectFunction(eval("field"&i), r(eval("field"&i)))
	end if%> 
	<%if eval("ftype"&i) = "textarea" then%>
		<textarea rows=10 cols=35 ID=textarea NAME="<%=eval("field"&i)%>"><%=r(eval("field"&i))%></textarea>
	<%end if%>
	<%if eval("ftype"&i) = "checkbox" then%>
		<input type=checkbox <%if r(eval("field"&i)) then%>checked <%end if%>ID="Text2" NAME="<%=eval("field"&i)%>">
	<%end if
    if eval("ftype"&i) = "browse" then
    sql2 = "select * from pics where num="&request("num")
    r2.open sql2,strconn,1,3
    if not r2.eof then theFolder=r2("pic_page")
    if eval("field"&i)="pic_name" then thefolder=thefolder&"/JPEG"
    r2.close
	filepath = ""
	%>
		<input type=text NAME="<%=eval("field"&i)%>" readonly value="<%=r(eval("field"&i))%>" ID="File1">&nbsp;&nbsp;<input value="דפדף..." type=button onclick="window.open('upload.asp?fieldname=<%=eval("field"&i)%>&thefolder=<%=theFolder%>','','width=320, height=120')">&nbsp;
        <img height=150 align=middle src="../gallery/<%=theFolder%>/<%=r(eval("field"&i))%>" width=200/>
	<%end if%>
	</td><td valign=top><%=oblig%><%=eval("fname"&i)%></td></tr>
<%next
%>
<tr><td colspan=2 align=center><input style="width:50px" class="submit" type="button" onclick="history.back();history.back();" value="חזור" name=sumbitButton ID="Submit1">  <input style="width:50px" onclick="delete_sub()" class="submit" type="button" value="מחק" name=deleteButton ID="Submit2">  <input style="width:50px" class="submit" type="submit" value="עדכן" name=sumbitButton ID="Submit3"></td></tr>
</table>
</form>
</td></tr></table>
</td><td width=1 bgcolor="#990000"></td></tr><tr align=right><td height=1 class=text_white colspan=10 bgcolor=#990000 align=center></td></tr></table>
</td></tr></table>