
<!-- #include file="secure.inc" -->
<%
set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")
set r2=Server.CreateObject("ADODB.Recordset")
	indexField = ""	
	subTable = ""

thetable = request("page")
currpage = thetable
sql = "SELECT * FROM "&thetable
r.open sql,strconn,1,3
allFields = r.Fields.Count

%>
<!-- #include file="menu.inc" -->
<script>
	function refreshPage(x){
		location.href="add.asp?page=<%=request("page")%>&itemType="+x+"";
	}
</script>
<%

function twoDigits(x)
	if x<10 then 
		twoDigits = "0"&x
	else
		twoDigits = x
	end if
end function

if request("sumbitButton")<>"" then
if eval("isAuto")="yes" then
    fields = field1
    if eval("ftype1")="number" then
        fields_content = request(field1)
    else
        fields_content = "'"&replace(request(field1),"'","&#8217;")&"'"
    end if
else
    fields = eval("indexField")&","&field1
    sql1 = "select * from "&theTable&" order by "&indexField&" desc"
    r1.open sql1,strconn,1,3
    if not r1.eof then 
        x = r1("recId")+1
    else
        x = 1
    end if
    if eval("ftype1")="number" then
        fields_content = eval("indexField")&","&request(field1)
    else
        fields_content = x&",'"&replace(request(field1),"'","&#8217;")&"'"
    end if
    r1.close
end if

for i = 2 to numOfFields
	if eval("field"&i)<>"" then
		curr_fields = replace(request(eval("field"&i)),"'","&#8217;")
		if (eval("ftype"&i)="number" or eval("ftype"&i)="select_num" or eval("ftype"&i)="date") and curr_fields="" then
			fields = fields
		else
			fields = fields&" ,"&eval("field"&i)
			if eval("ftype"&i) = "checkbox" then
				if curr_fields = "on" then
					curr_fields = -1
				else
					curr_fields = 0
				end if
			else
				if eval("ftype"&i) = "text" or eval("ftype"&i) = "textarea" or eval("ftype"&i) = "select_text" or eval("ftype"&i) = "browse" or eval("ftype"&i) = "color"  then
				    curr_fields = "'"&curr_fields&"'"
				else 
					if eval("ftype"&i) = "date" then
						curr_fields = "'"&year(curr_fields)&"-"&twoDigits(month(curr_fields))&"-"&twoDigits(day(curr_fields))&"'"
					end if
				end if
			end if
			fields_content = fields_content&" ,"&curr_fields
		end if
	end if
next
sql = "INSERT INTO "&thetable&" ("&fields&") VALUES ("&fields_content&")"
r1.open sql,strconn,1,3
'response.write sql
'response.write err.description
'response.end
response.redirect "inner.asp?page="&request("page")
end if

sub selectFunction(theField)
    select case theField
        case "orgtype":
	        sql = "SELECT * FROM [imknasim].[dbo].[orgtype] order by theName"
	        r2.open sql,strconn,1,3
		    response.Write "<select style='width:100%;height:30px;border:0px;background-color:#EAEAEA;border-radius:8px;box-sizing:border-box;padding-right:10px;font-size:18px' name=orgtype><option value=0>בחר...</option>"
		    while not r2.EOF
		        response.Write "<option value="&r2("theIndex")&">"&r2("theName")&"</option>"
		        r2.MoveNext
		    wend
            r2.close
		    response.Write "</select>"
		case "fieldtype":
	        sql = "SELECT * FROM [imknasim].[dbo].[fieldtype] order by theName"
	        r2.open sql,strconn,1,3
		    response.Write "<select style='width:100%;height:30px;border:0px;background-color:#EAEAEA;border-radius:8px;box-sizing:border-box;padding-right:10px;font-size:18px' name=fieldtype><option value=0>בחר...</option>"
		    while not r2.EOF
		        response.Write "<option value="&r2("theIndex")&">"&r2("theName")&"</option>"
		        r2.MoveNext
		    wend
            r2.close
		    response.Write "</select>"
		case "fieldsection":
	        sql = "SELECT * FROM [imknasim].[dbo].[section] order by theName"
	        r2.open sql,strconn,1,3
		    response.Write "<select style='width:100%;height:30px;border:0px;background-color:#EAEAEA;border-radius:8px;box-sizing:border-box;padding-right:10px;font-size:18px' name=fieldsection><option value=0>בחר...</option>"
		    while not r2.EOF
		        response.Write "<option value="&r2("theIndex")&">"&r2("theName")&"</option>"
		        r2.MoveNext
		    wend
            r2.close
		    response.Write "</select>"
		case "choicetable":
	        sql = "SELECT * FROM [imknasim].[dbo].[choicetable] order by theName"
	        r2.open sql,strconn,1,3
		    response.Write "<select style='width:100%;height:30px;border:0px;background-color:#EAEAEA;border-radius:8px;box-sizing:border-box;padding-right:10px;font-size:18px' name=choicetable><option value=0>בחר...</option>"
		    while not r2.EOF
		        response.Write "<option value="&r2("theIndex")&">"&r2("theName")&"</option>"
		        r2.MoveNext
		    wend
            r2.close
		    response.Write "</select>"
		case "validation":
	        sql = "SELECT * FROM [imknasim].[dbo].[validationtype] order by theName"
	        r2.open sql,strconn,1,3
		    response.Write "<select style='width:100%;height:30px;border:0px;background-color:#EAEAEA;border-radius:8px;box-sizing:border-box;padding-right:10px;font-size:18px' name=validation><option value=0>בחר...</option>"
		    while not r2.EOF
		        response.Write "<option value="&r2("theIndex")&">"&r2("theName")&"</option>"
		        r2.MoveNext
		    wend
            r2.close
		    response.Write "</select>"
	    
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
 ###   indexField = "NAME OF KEY FIELD OF THE TABLE"
 ###   fname1, 2, 3... = "TEXT TO PRESENT FOR THE FIELD"
 ###   field1, 2, 3... = "THE NAME OF THE FIELD IN THE TABLE"
 ###   ftype1, 2, 3... = "THE FIELD'S TYPE"
 ###   subTable = "IF THERE IS A SUB TABLE TO CONECT - WRITE THE NAME OF THE SUB TABLE"
 ###   subIndex = "IF THERE IS A SUB TABLE TO CONECT - THE NAME OF THE KEY FIELD OF THE SUB TABLE"
 ###   subWhere = "IF THERE IS A SUB TABLE TO CONECT - THE NAME OF THE FILED OF THE PARENT TABLE THAT CONACTS TO THE SUB TABLE"
 ###   subField = "IF THERE IS A SUB TABLE TO CONECT - THE TITLE FIELD OF THE SUB FIELD TO SHOW ON THE PAGE"
 ###
 ###   RULES:
 ###   NAMES OF TABLES AND FIELDS CAN'T BE: page, table, where, field, text, date, number, on ,off, name, add, edit
 ###   TYPES OF FIELDS: text; number; date; checkbox; select; *browse;
 ###   * browse field calls the upload.asp and upload1.asp files, should be adjusted per server
 ###
 #################################################################################################################################
------------------------------------------------------------------------------------------------------------------------------------>

<html><head><title>מערכת ניהול</title>
<meta charset="UTF-8">
<!-- #INCLUDE file="FCKeditor/fckeditor.asp" -->
<script language=javascript>
function checkForm(theForm)
{
	if(theForm.<%=field1%>.value=="")
	{
		//document.getElementById("errorTr").style.display="inline";
		return true;
	}
	else
	{
		document.getElementById("errorTr").style.display="none";
		return true;
	}
}
</script>
<link rel="stylesheet" type="text/css" href="dhtmlgoodies_calendar.css">
<script language=javascript src="dhtmlgoodies_calendar.js"></script>
<link rel="stylesheet" type="text/css" href="admin.css">
</head>
<body style="margin:0px;padding:0px;direction:rtl;text-align:right">
<div style="float:right;width:100%">
<div style="float:right;width:100%;text-align:center;font-size:40px;background-color:#FAFAFA;height:50px">מערכת ניהול</div>
<div style="float:right;width:100%;height:100%">
<div style="float:right;width:20%;height:1000px;background-color:#EAEAEA">
    <!-- #include file="toolbar.inc" -->
</div>
<div style="float:right;width:50%;box-sizing:border-box;padding:10px">
<form name=oform method=post onsubmit="return checkForm(this)">
<%for i=1 to numOfFields
	toContinue=true
	'if request("page")="[stamps].[dbo].[MainTable]" and request("itemType")<>"" then
	'	sql1 = "select * from [stamps].[dbo].[fieldItem] where theField="&request("itemType")&" and theItem="&i&""
		'response.write sql1&"<br>"
	'	r1.open sql1,strconn,1,3
	'	if r1.eof then toContinue=false
	'	r1.close
	'end if
	if toContinue=true then
		if i = 1 then
			oblig = "<span class='oblig'>*</span>"
		else
			oblig = ""
		end if
		%>
		<div style="float:right;width:100%;margin-bottom:10px">
		<div style='float:right;width:20%;font-size:24px'><%=oblig%><%=eval("fname"&i)%>:</div>
		<div style="float:right;width:75%">
		<%if eval("ftype"&i) = "color" then%>
		<nobr><input type="color" ID="Text1" NAME="<%=eval("field"&i)%>" value="<%=r(eval("field"&i))%>">
		<%end if%>
		<%if eval("ftype"&i) = "date" then%>
			<nobr><input readonly onclick="displayCalendar(this,'dd/mm/yyyy',this)" value="" type=text ID="Text3" NAME="<%=eval("field"&i)%>"> <input class="submit" type="button" style="width:70" value="תאריכון" onclick="    displayCalendar(document.forms[0].<%=eval("field"&i)%>,'dd/mm/yyyy',this)" ID="Button1" NAME="Button1"></nobr>
		<%end if
		if eval("ftype"&i) = "text" or eval("ftype"&i) = "number" then
			theValue=""
			if request("page")="[stamps].[dbo].[articles]" and eval("field"&i)="oldIndex" then
				sql1 = "select * from [stamps].[dbo].[articles] order by oldIndex desc"
				r1.open sql1,strconn,1,3
				if not r1.eof then theValue=r1("oldIndex")+1
				r1.close
			end if
			%>
			<input style="width:100%;height:30px;border:0px;background-color:#EAEAEA;border-radius:8px;box-sizing:border-box;padding-right:10px;font-size:18px" type=text <%if eval("ftype"&i) = "number" then%> onkeypress="if((event.keyCode<48 || event.keyCode>57) && event.keyCode!=46) return false" <%end if%> value="<%=thevalue%>" ID="Text1" NAME="<%=eval("field"&i)%>">
		<%end if
		if eval("ftype"&i) = "browse" then
		filepath = ""
		%>
			<input style="width:70%;height:30px;border:0px;background-color:#EAEAEA;border-radius:8px;box-sizing:border-box;padding-right:10px;font-size:18px;margin-left:2%" type=text NAME="<%=eval("field"&i)%>" readonly value="" ID="File1" /><input style="width:28%;height:30px;border:0px;background-color:#EAEAEA;border-radius:8px;box-sizing:border-box;padding-right:10px;font-size:18px" value="דפדף..." type=button onclick="window.open('upload.asp?fieldname=<%=eval("field"&i)%>','','width=480, height=240')">
		<%end if
		if eval("ftype"&i) = "textarea" then
			Dim NewsEdit
			Set NewsEdit = New FCKeditor
			NewsEdit.BasePath = "FCKeditor/"
			NewsEdit.Config("AutoDetectLanguage") = False
			if eval("lang"&i)="he" then
				NewsEdit.Config("DefaultLanguage") = "he"
				NewsEdit.Config("ContentLangDirection") = "rtl"
			else
				NewsEdit.Config("DefaultLanguage") = "en"
				NewsEdit.Config("ContentLangDirection") = "ltr"
			end if
			NewsEdit.Config("CustomConfigurationsPath") = "FCKeditor/adminconfig.js"
			NewsEdit.Value = ""
			NewsEdit.Height = 450
			NewsEdit.Width = 580
			NewsEdit.Create eval("field"&i)
		end if%>
		<%if eval("ftype"&i) = "checkbox" then%>
			<input type=checkbox ID="Text2" NAME="<%=eval("field"&i)%>">
		<%end if%> 
		<%if eval("ftype"&i) = "select_text" or eval("ftype"&i) = "select_num" then
			selectFunction(eval("field"&i))
		end if%> 
		</div>
		</div>
	<%
		end if
	next
%>
<div style="float:right;width:100%">
    <input class="submit" style="width:200px;height:40px;background-color:#EAEAEA;box-sizing:border-box;border-radius:8px;border:0px;margin-top:20px;font-size:24px" type="submit" value="עדכן" name=sumbitButton ID="Submit2">&nbsp;
    
</div>
</form>
</div>


