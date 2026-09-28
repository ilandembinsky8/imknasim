<html>
  <head>
    <meta charset="UTF-8">
      <link rel="stylesheet" type="text/css" href="admin.css">
</head>
  <body>
    <%
' ***************************************************************

    ' changes: number of pages, pages names and titles

    ' ***************************************************************
    set rs_tb=Server.CreateObject("ADODB.Recordset")
    pages = 25 'Number of tables to edit

	page1 = "[imknasim].[dbo].[conference]"
    pName1 = "כנסים"
    pTitle1 = "thename"
    pIndex1 = "theindex"

	page2 = "[imknasim].[dbo].[itemtype]"
    pName2 = "סוג פריט"
    pTitle2 = "thename"
    pIndex2 = "theIndex"

	page3 = "[imknasim].[dbo].[organization]"
    pName3 = "ארגון"
    pTitle3 = "thename"
    pIndex3 = "theIndex"

	page4 = "[imknasim].[dbo].[orgtype]"
    pName4 = "סוג ארגון"
    pTitle4 = "thename"
    pIndex4 = "theindex"

	page5 = "[imknasim].[dbo].[sponsor]"
    pName5 = "נותני חסות"
    pTitle5 = "thename"
    pIndex5 = "theIndex"

	page6 = "[imknasim].[dbo].[title]"
    pName6 = "תואר"
    pTitle6 = "thename"
    pIndex6 = "theIndex"

    page11 = "[imknasim].[dbo].[section]"
    pName11 = "תחום"
    pTitle11 = "thename"
    pIndex11 = "theIndex"

    page12 = "[imknasim].[dbo].[fieldtype]"
    pName12 = "סוג שדה"
    pTitle12 = "thename"
    pIndex12 = "theIndex"

    page13 = "[imknasim].[dbo].[choicetable]"
    pName13 = "טבלאות בחירה"
    pTitle13 = "thename"
    pIndex13 = "theIndex"

    page14 = "[imknasim].[dbo].[sex]"
    pName14 = "מין"
    pTitle14 = "thename"
    pIndex14 = "theIndex"

    page15 = "[imknasim].[dbo].[country]"
    pName15 = "מדינה"
    pTitle15 = "thename"
    pIndex15 = "theIndex"

    page16 = "[imknasim].[dbo].[validationtype]"
    pName16 = "ולידציה"
    pTitle16 = "thename"
    pIndex16 = "theIndex"

    page17 = "[imknasim].[dbo].[profession]"
    pName17 = "מקצועות"
    pTitle17 = "thename"
    pIndex17 = "theIndex"

    page18 = "[imknasim].[dbo].[participanttype]"
    pName18 = "סוג משתתף"
    pTitle18 = "thename"
    pIndex18 = "theIndex"

    page19 = "[imknasim].[dbo].[participation]"
    pName19 = "סוג השתתפות"
    pTitle19 = "thename"
    pIndex19 = "theIndex"

    page20 = "[imknasim].[dbo].[registertype]"
    pName20 = "סוג הרשמה"
    pTitle20 = "thename"
    pIndex20 = "theIndex"

    page21 = "[imknasim].[dbo].[paymenttype]"
    pName21 = "סוג תשלום"
    pTitle21 = "thename"
    pIndex21 = "theIndex"

    page22 = "[imknasim].[dbo].[paymentstatus]"
    pName22 = "סטטוס תשלום"
    pTitle22 = "thename"
    pIndex22 = "theIndex"

    page22 = "[imknasim].[dbo].[roomtype]"
    pName22 = "סוג חדר"
    pTitle22 = "thename"
    pIndex22 = "theIndex"

    page23 = "[imknasim].[dbo].[nutrition]"
    pName23 = "העדפות מזון"
    pTitle23 = "thename"
    pIndex23 = "theIndex"

    page24 = "[imknasim].[dbo].[accesstype]"
    pName24 = "סוגי נגישות"
    pTitle24 = "thename"
    pIndex24 = "theIndex"

    page25 = "[imknasim].[dbo].[documents]"
    pName25 = "מסמכים"
    pTitle25 = "thename"
    pIndex25 = "theIndex"

    page8 = "[imknasim].[dbo].[fields]"
    pName8 = "שדות"
    pTitle8 = "thename"
    pIndex8 = "theIndex"

	page = currpage
    %>

    <div style="float:right;width:100%;height:200%;font-size:28px;line-height:135%;box-sizing:border-box;padding:10px">
      <%for i=1 to pages%>
        <a
		  
          <%if page=eval("page"&i) then%>style="font-weight:bold;font-size:28px"<%end if%>
            href="inner.asp?page=<%=eval("page"&i)%>" style="font-size:28px"><%=eval("pName"&i)%>
        </a>
        <br>
          <%next%>
     </div>
  </body>
</html>