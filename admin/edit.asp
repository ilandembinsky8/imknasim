<!-- #include file="secure.inc" -->
<%
set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")
set r2=Server.CreateObject("ADODB.Recordset")
	indexField = ""
	subTable = ""

thetable = request("page")
currpage = thetable
allFields = request("all")

%>
<!-- #include file="menu.inc" -->
<script>
	function refreshPage(x){
		location.href="edit.asp?page=<%=request("page")%>&num=<%=request("num")%>&itemType="+x+"";
	}

	function gotoPage(x){
		location.href=x;
	}
</script>
<script>
function createPalette(image, options = {}) {
  const {
    sampleSize = 100,
    colorBucket = 32
  } = options;

  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d", { willReadFrequently: true });

  const scale = Math.min(
    sampleSize / image.width,
    sampleSize / image.height,
    1
  );

  canvas.width = Math.max(1, Math.floor(image.width * scale));
  canvas.height = Math.max(1, Math.floor(image.height * scale));

  ctx.drawImage(image, 0, 0, canvas.width, canvas.height);

  const { data } = ctx.getImageData(
    0,
    0,
    canvas.width,
    canvas.height
  );

  // --------------------------------------------------
  // 1. Quantize colors and count popularity
  // --------------------------------------------------

  const colors = new Map();

  for (let i = 0; i < data.length; i += 4) {
    const r = data[i];
    const g = data[i + 1];
    const b = data[i + 2];
    const a = data[i + 3];

    if (a < 128) continue;

    // Ignore nearly-white background pixels
    if (r > 245 && g > 245 && b > 245) continue;

    const qr = Math.floor(r / colorBucket) * colorBucket;
    const qg = Math.floor(g / colorBucket) * colorBucket;
    const qb = Math.floor(b / colorBucket) * colorBucket;

    const key = `${qr},${qg},${qb}`;

    colors.set(key, (colors.get(key) || 0) + 1);
  }

  const dominant = [...colors.entries()]
    .sort((a, b) => b[1] - a[1])
    .map(([key]) => key.split(",").map(Number));

  // --------------------------------------------------
  // 2. Select 6 distinct palette colors from image
  // --------------------------------------------------

  const paletteColors = [];

  for (const color of dominant) {
    if (
      paletteColors.every(
        existing => colorDistance(color, existing) > 40
      )
    ) {
      paletteColors.push(color);
    }

    if (paletteColors.length === 6) break;
  }

  // Fallback if image doesn't have 6 distinct colors
  while (paletteColors.length < 6) {
    paletteColors.push([128, 128, 128]);
  }

  // --------------------------------------------------
  // 3. Generate 2 contrasting / accent colors
  //    (derived from the 2 most dominant palette colors)
  // --------------------------------------------------

  const contrastColors = [
    complementaryColor(paletteColors[0]),
    complementaryColor(paletteColors[1])
  ];

  // --------------------------------------------------
  // 4. Return total 8 colors:
  //    [0-5] = 6 Palette Colors
  //    [6-7] = 2 Contrast Colors
  // --------------------------------------------------

  return [
    ...paletteColors,
    ...contrastColors
  ].map(rgbToHex);
}

// ======================================================
// Helpers
// ======================================================

function colorDistance(a, b) {
  return Math.sqrt(
    Math.pow(a[0] - b[0], 2) +
    Math.pow(a[1] - b[1], 2) +
    Math.pow(a[2] - b[2], 2)
  );
}

function rgbToHex([r, g, b]) {
  return (
    "#" +
    [r, g, b]
      .map(v => Math.max(0, Math.min(255, Math.round(v)))
        .toString(16)
        .padStart(2, "0")
      )
      .join("")
  );
}

// RGB -> HSL
function rgbToHsl([r, g, b]) {
  r /= 255;
  g /= 255;
  b /= 255;

  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);

  let h;
  let s;
  const l = (max + min) / 2;

  if (max === min) {
    h = s = 0;
  } else {
    const d = max - min;

    s = l > 0.5
      ? d / (2 - max - min)
      : d / (max + min);

    switch (max) {
      case r:
        h = (g - b) / d + (g < b ? 6 : 0);
        break;

      case g:
        h = (b - r) / d + 2;
        break;

      case b:
        h = (r - g) / d + 4;
        break;
    }

    h /= 6;
  }

  return [h * 360, s, l];
}

// HSL -> RGB
function hslToRgb(h, s, l) {
  h /= 360;

  let r, g, b;

  if (s === 0) {
    r = g = b = l;
  } else {
    const hue = (p, q, t) => {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;

      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;

      return p;
    };

    const q =
      l < 0.5
        ? l * (1 + s)
        : l + s - l * s;

    const p = 2 * l - q;

    r = hue(p, q, h + 1 / 3);
    g = hue(p, q, h);
    b = hue(p, q, h - 1 / 3);
  }

  return [
    Math.round(r * 255),
    Math.round(g * 255),
    Math.round(b * 255)
  ];
}

// Generate a high-contrast complementary color
function complementaryColor(rgb) {
  const [h, s, l] = rgbToHsl(rgb);

  // Invert hue by 180° and adjust lightness/saturation for optimum visibility
  const newL = l > 0.5 ? Math.max(0.2, l - 0.4) : Math.min(0.8, l + 0.4);

  return hslToRgb(
    (h + 180) % 360,
    Math.min(1, Math.max(0.5, s * 1.2)),
    newL
  );
}
</script>
<%

if request("del") = 1 then
	sql = "DELETE FROM "&thetable&" WHERE "&indexField&" = "&request("num")
	'response.write sql
	'response.end
	r1.open sql,strconn,1,3
	'response.Write sql
	response.Redirect "inner.asp?page="&request("page")
end if
if subTable<>"" then
sql = "select * from pages where num="&request("num")
r.open sql,strconn,1,3
theName=r("page_title")
r.close
sql = "SELECT * FROM "&subTable&" WHERE "&subWhere&" = '"&theName&"'"
'response.write sql
r1.open sql,strconn,1,3
end if
if request("sumbitButton")<>"" then
sql = "UPDATE "&thetable&" set "
for i=1 to numOfFields
	
	if eval("field"&i)<>"" then
		if i>1 then
			sql = sql&" ,"
		end if
		tablevalue = replace(request(eval("field"&i)),"'","&#8217;")
		if tablevalue = "" and (eval("ftype"&i) = "number" or eval("ftype"&i) = "select_num") then
			tablevalue = "0"
		end if
		if tablevalue = "" and (eval("ftype"&i) = "text" or eval("ftype"&i) = "select_text") then
			tablevalue = " "
		end if
		if tablevalue = "" and eval("ftype"&i) = "date" then
			tablevalue = "1/1/1"
		end if
		if eval("ftype"&i) = "checkbox" then
			if tablevalue = "on" then
				tablevalue = -1
			else
				tablevalue = 0
			end if
		else
		    if eval("ftype"&i)="date" then
				if day(tablevalue)>12 then
            	    d = day(tablevalue)
				    m = month(tablevalue)
				    y = year(tablevalue)
                   tablevalue = "'"&y&"-"&m&"-"&d&"'"
                else
                    d = day(tablevalue)
				    m = month(tablevalue)
				    y = year(tablevalue)
				    tablevalue = "'"&dateserial(y,d,m)&"'" 
			    end if
			else
                tablevalue = "'"&tablevalue&"'"
		    end if
		end if
		sql = sql&eval("field"&i)&" = "&tablevalue
	end if
next
sql = sql &" WHERE "&indexField&" = "&request("num")
'response.write sql
'response.end
r.open sql,strconn,1,3
end if

'if thetable="[stamps].[dbo].[MainTable]" then
	'if request("itemType")&"a"="a" then
	'	sql = "select * from [stamps].[dbo].[MainTable] where theIndex="&request("num")
	'	r.open sql,strconn,1,3
	'	if not r.eof then
	'		theItem=r("itemType")
	'		r.close
	'		if theItem&"a"="a" then theItem=1
	'		response.redirect "edit.asp?page="&request("page")&"&num="&request("num")&"&itemType="&theItem
	'	else
	'		r.close
	'	end if
	'end if
'end if

sql = "SELECT * FROM "&thetable
'sql1 = sql &" WHERE "&indexField&" < "&request("num")& " order by theIndex desc"
'sql2 = sql &" WHERE "&indexField&" > "&request("num")& " order by theIndex"
sql = sql & " WHERE "&indexField&" = "&request("num")& " order by "&field1&" desc"
'response.Write sql
'response.End
r.Open sql,strconn,1,3

sub selectFunction(theField, fieldValue)
	select case theField
        case "orgtype":
	        sql = "SELECT * FROM [imknasim].[dbo].[orgtype] order by theName"
	        r2.open sql,strconn,1,3
		    response.Write "<select style='width:100%;height:30px;border:0px;background-color:#EAEAEA;border-radius:8px;box-sizing:border-box;padding-right:10px;font-size:18px' name=orgtype><option value=0>בחר...</option>"
		    while not r2.EOF
		        if r2("theIndex") = fieldValue then
		            response.Write "<option selected value="&r2("theIndex")&">"&r2("theName")&"</option>"
		        else
		            response.Write "<option value="&r2("theIndex")&">"&r2("theName")&"</option>"
		        end if
		        r2.MoveNext
		    wend
            r2.close
		    response.Write "</select>"
		case "fieldtype":
	        sql = "SELECT * FROM [imknasim].[dbo].[fieldtype] order by theName"
	        r2.open sql,strconn,1,3
		    response.Write "<select style='width:100%;height:30px;border:0px;background-color:#EAEAEA;border-radius:8px;box-sizing:border-box;padding-right:10px;font-size:18px' name=fieldtype><option value=0>בחר...</option>"
		    while not r2.EOF
		        if r2("theIndex") = fieldValue then
		            response.Write "<option selected value="&r2("theIndex")&">"&r2("theName")&"</option>"
		        else
		            response.Write "<option value="&r2("theIndex")&">"&r2("theName")&"</option>"
		        end if
		        r2.MoveNext
		    wend
            r2.close
		    response.Write "</select>"
		case "fieldsection":
	        sql = "SELECT * FROM [imknasim].[dbo].[section] order by theName"
	        r2.open sql,strconn,1,3
		    response.Write "<select style='width:100%;height:30px;border:0px;background-color:#EAEAEA;border-radius:8px;box-sizing:border-box;padding-right:10px;font-size:18px' name=fieldsection><option value=0>בחר...</option>"
		    while not r2.EOF
		        if r2("theIndex") = fieldValue then
		            response.Write "<option selected value="&r2("theIndex")&">"&r2("theName")&"</option>"
		        else
		            response.Write "<option value="&r2("theIndex")&">"&r2("theName")&"</option>"
		        end if
		        r2.MoveNext
		    wend
            r2.close
		    response.Write "</select>"
		case "choicetable":
	        sql = "SELECT * FROM [imknasim].[dbo].[choicetable] order by theName"
	        r2.open sql,strconn,1,3
		    response.Write "<select style='width:100%;height:30px;border:0px;background-color:#EAEAEA;border-radius:8px;box-sizing:border-box;padding-right:10px;font-size:18px' name=choicetable><option value=0>בחר...</option>"
		    while not r2.EOF
		        if r2("theIndex") = fieldValue then
		            response.Write "<option selected value="&r2("theIndex")&">"&r2("theName")&"</option>"
		        else
		            response.Write "<option value="&r2("theIndex")&">"&r2("theName")&"</option>"
		        end if
		        r2.MoveNext
		    wend
            r2.close
		    response.Write "</select>"
		case "validation":
	        sql = "SELECT * FROM [imknasim].[dbo].[validationtype] order by theName"
	        r2.open sql,strconn,1,3
		    response.Write "<select style='width:100%;height:30px;border:0px;background-color:#EAEAEA;border-radius:8px;box-sizing:border-box;padding-right:10px;font-size:18px' name=validation><option value=0>בחר...</option>"
		    while not r2.EOF
		        if r2("theIndex") = fieldValue then
		            response.Write "<option selected value="&r2("theIndex")&">"&r2("theName")&"</option>"
		        else
		            response.Write "<option value="&r2("theIndex")&">"&r2("theName")&"</option>"
		        end if
		        r2.MoveNext
		    wend
            r2.close
		    response.Write "</select>"
	    
	end select
end sub
%>

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
<link rel="stylesheet" type="text/css" href="dhtmlgoodies_calendar.css">
<link rel="stylesheet" type="text/css" href="admin.css">
<script language=javascript src="dhtmlgoodies_calendar.js"></script>
</head>
<body style="margin:0px;padding:0px;direction:rtl;text-align:right">
<div style="float:right;width:100%">
<div style="float:right;width:100%;text-align:center;font-size:40px;background-color:#FAFAFA;height:50px">מערכת ניהול</div>
<div style="float:right;width:100%;height:100%">
<div style="float:right;width:20%;height:1000px;background-color:#EAEAEA">
    <!-- #include file="toolbar.inc" -->
</div>
<div style="float:right;width:50%;box-sizing:border-box;padding:10px">
<% 
if request("page")="[stamps].[dbo].[MainTable]" then
%>
<div style="float:right;width:100%;font-size:24px;margin-bottom:20px"><a style="font-size:24px" href="http://stamps.nez.co.il/stamp.asp?item=<%=request("num")%>" target="_blank">צפייה בעמוד</a></div>
<% 
end if
%>
<form name=oform method=post onsubmit="return checkForm(this)">
<%
	for i=1 to numOfFields
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
		<nobr><input type="color" ID="color<%=i%>" NAME="<%=eval("field"&i)%>" value="<%=r(eval("field"&i))%>">&nbsp;<span onclick="pasteColor(<%=i%>)" style="cursor:pointer">הדבק צבע</span>
		<%end if%>
		<%if eval("ftype"&i) = "date" then%>
			<div style="float:right;margin-left:4%">
			<% 
			if year(r(eval("field"&i)))=2001 and month(r(eval("field"&i)))=1 and day(r(eval("field"&i)))=1 then
			%>
			<nobr><input readonly onclick="displayCalendar(this,'dd/mm/yyyy',this)" type=text id="<%=eval("field"&i)%>" NAME="<%=eval("field"&i)%>"> <input class="submit" type="button" style="width:70" value="תאריכון" onclick="displayCalendar(document.forms[0].<%=eval("field"&i)%>,'dd/mm/yyyy',this)" ID="Button1" NAME="Button1"></nobr>
			<% 
			else
			%>
			<nobr><input readonly onclick="displayCalendar(this,'dd/mm/yyyy',this)" value="<%=r(eval("field"&i))%>" type=text id="<%=eval("field"&i)%>" NAME="<%=eval("field"&i)%>"> <input class="submit" type="button" style="width:70" value="תאריכון" onclick="displayCalendar(document.forms[0].<%=eval("field"&i)%>,'dd/mm/yyyy',this)" ID="Button1" NAME="Button1"></nobr>
			<% 
			end if
			%>
			</div>
			<div style="width:25%;float:right;font-size:20px;cursor:pointer" onclick="document.getElementById('<%=eval("field"&i)%>').value=''">נקה</div>
		<%end if
		if eval("ftype"&i) = "text" or eval("ftype"&i) = "number" then
		thevalue = r(eval("field"&i))
			if thevalue<>"" then
				thevalue = replace(thevalue,"""","&quot;")
			end if%>
			<input style="width:100%;height:30px;border:0px;background-color:#EAEAEA;border-radius:8px;box-sizing:border-box;padding-right:10px;font-size:18px" type=text <%if eval("ftype"&i) = "number" then%> onkeypress="if((event.keyCode<48 || event.keyCode>57) && event.keyCode!=46) return false" <%end if%> value="<%=thevalue%>" ID="Text1" NAME="<%=eval("field"&i)%>">
		<%end if
		if eval("ftype"&i) = "browse" then
		filepath = ""
		%>
			<div style="float:right;width:70%;margin-left:4%">
			<input style="width:100%;height:30px;border:0px;background-color:#EAEAEA;border-radius:8px;box-sizing:border-box;padding-right:10px;font-size:18px" type=text ID="<%=eval("field"&i)%>" NAME="<%=eval("field"&i)%>" readonly value="<%=r(eval("field"&i))%>" ID="File1">&nbsp;&nbsp;<input value="דפדף..." type=button onclick="window.open('upload.asp?fieldname=<%=eval("field"&i)%>','','width=480, height=240')">&nbsp;
			</div>
			<div style="width:25%;float:right;font-size:20px;cursor:pointer" onclick="document.getElementById('<%=eval("field"&i)%>').value=''">נקה</div>
			<div style="width:100%;float:right">
			<img height=150 align=middle src="../images/<%=r(eval("field"&i))%>" width=200/>
			</div>
		<%end if
		if eval("ftype"&i) = "textarea" then
			Dim NewsEdit
			Set NewsEdit = New FCKeditor
			NewsEdit.BasePath = "FCKeditor/"
			NewsEdit.Config("AutoDetectLanguage") = False
			if request("page")="[stamps].[dbo].[pages]" then
				if r("hebrew")=0 then
					lang3="en"
				else
					lang3="he"
				end if
			end if
			if eval("lang"&i)="he" then
				NewsEdit.Config("DefaultLanguage") = "he"
				NewsEdit.Config("ContentLangDirection") = "rtl"
			else
				NewsEdit.Config("DefaultLanguage") = "en"
				NewsEdit.Config("ContentLangDirection") = "ltr"
			end if
			NewsEdit.Config("CustomConfigurationsPath") = "FCKeditor/adminconfig.js"
			NewsEdit.Value = r(eval("field"&i))
			NewsEdit.Height = 450
			NewsEdit.Width = 580
			NewsEdit.Create eval("field"&i)
		end if%>
		<%if eval("ftype"&i) = "checkbox" then%>
			<input type=checkbox <%if r(eval("field"&i)) then%>checked <%end if%>ID="Text2" NAME="<%=eval("field"&i)%>">
		<%end if%> 
		<%if eval("ftype"&i) = "select_text" or eval("ftype"&i) = "select_num" then
			call selectFunction(eval("field"&i), r(eval("field"&i)))
		end if%> 
		</div>
		</div>
	<%
	end if
next
%>
<div style="float:right;width:100%">
    <input class="submit" style="width:200px;height:40px;background-color:#EAEAEA;box-sizing:border-box;border-radius:8px;border:0px;margin-top:20px;font-size:24px" type="submit" value="עדכן" name=sumbitButton ID="Submit2">&nbsp;
    <input class="submit" style="width:200px;height:40px;background-color:#EAEAEA;box-sizing:border-box;border-radius:8px;border:0px;margin-top:20px;font-size:24px" type=button value="מחק" onclick="if (confirm('למחוק?')) gotoPage('edit.asp?page=<%=request("page")%>&num=<%=request("num")%>&del=1')" ID="Button2" NAME="Button2">
</div>
</form>
</div>
<%
if request("page")="[imknasim].[dbo].[conference]" then
	%>
	<div style="float:right;width:20%">
		<div style="float:right;width:100%">
		<img id="source" src="../images/<%=r(eval("field2"))%>" style="max-width:100%" crossorigin="anonymous">
		</div>
		

		<!-- 6 Palette Swatches -->
		<div style="float:right;width:100%;font-size:1.5rem">צבעי פאלטה</div>
		<div style="float:right;width:100%">
			<div id="c1" style="float:left; width:30%; margin-right:3%;height:100px;text-align:center;padding-top:10px;margin-bottom:10px;border-radius:10px;cursor:pointer" onclick="copyColor(this)"></div>
			<div id="c2" style="float:left; width:30%; margin-right:3%;height:100px;text-align:center;padding-top:10px;margin-bottom:10px;border-radius:10px;cursor:pointer" onclick="copyColor(this)"></div>
			<div id="c3" style="float:left; width:30%; margin-right:3%;height:100px;text-align:center;padding-top:10px;margin-bottom:10px;border-radius:10px;cursor:pointer" onclick="copyColor(this)"></div>
			<div id="c4" style="float:left; width:30%; margin-right:3%;height:100px;text-align:center;padding-top:10px;margin-bottom:10px;border-radius:10px;cursor:pointer" onclick="copyColor(this)"></div>
			<div id="c5" style="float:left; width:30%; margin-right:3%;height:100px;text-align:center;padding-top:10px;margin-bottom:10px;border-radius:10px;cursor:pointer" onclick="copyColor(this)"></div>
			<div id="c6" style="float:left; width:30%; margin-right:3%;height:100px;text-align:center;padding-top:10px;margin-bottom:10px;border-radius:10px;cursor:pointer" onclick="copyColor(this)"></div>
		</div>
		<!-- 2 Contrast Swatches -->
		<div style="float:right;width:100%;font-size:1.5rem">צבעים משלימים</div>
		<div style="float:right;width:100%">
			<div id="c7" style="float:left; width:30%;margin-right:3%; height:100px;text-align:center;padding-top:10px;margin-bottom:10px;border-radius:10px;cursor:pointer" onclick="copyColor(this)"></div>
			<div id="c8" style="float:left; width:30%;margin-right:3%; height:100px;text-align:center;padding-top:10px;margin-bottom:10px;border-radius:10px;cursor:pointer" onclick="copyColor(this)"></div>
		</div>
		<input type="hidden" id="copied">
	</div>
	<script>
	function pasteColor(x){
		if (document.getElementById("copied").value=="") return;
		document.getElementById("color"+x+"").value = document.getElementById("copied").value;
		document.getElementById("copied").value = "";
	}
	function copyColor(element){
		const color = getComputedStyle(element).backgroundColor;
    	document.getElementById("copied").value = color;
	}
	const img = document.getElementById("source");

	img.onload = () => {
		const palette = createPalette(img);

		console.log("6 Palette Colors + 2 Contrast Colors:", palette);
		for (let i = 1; i <= palette.length; i++) {
			document.getElementById("c" + i).style.backgroundColor = palette[i - 1];
			document.getElementById("c" + i).innerHTML = palette[i - 1];
		}
	};

	
	</script>
	<%
end if
%>
<%if subTable<>"" then%>
<div style="float:right;width:20%">
<td valign=top align=right dir=rtl bgcolor="#dddddd">
<u>תתי נושא:</u><br><br>
<%
while not r1.EOF%>
<a href="edit_sub.asp?page=<%=subTable%>&num=<%=r1(subIndex)%>"><%=left(r1(subField),50)%><%if len(r1(subField))>39 then%>...<%end if%></a><br>
<%r1.MoveNext
wend%><br>
<input style="width:100px" type="submit" class="submit" value="הוסף תת נושא" name=sumbitButton ID="Submit1" onclick="location.reload('add_sub.asp?page=<%=subTable%>&num=<%=request("num")%>')">
</td><td width=10 bgcolor="#dddddd">&nbsp;</td>
</div><%
end if%>
</div>
