<link rel="stylesheet" type="text/css" href="admin.css">
<%
' ***************************************************************

' changes: number of pages, pages names and titles

' ***************************************************************
set rs_tb=Server.CreateObject("ADODB.Recordset")
pages = 4 'Number of tables to edit
page1 = "topics"
pName1 = "קטגוריות"
pTitle1 = "topic_title"
pIndex1 = "num"
page2 = "pages"
pName2 = "עמודים"
pTitle2 = "page_title"
pIndex2 = "num"
page3 = "pics"
pName3 = "תמונות"
pTitle3 = "pic_name"
pIndex3 = "num"
page4 = "info_pages"
pName4 = "עמודי תוכן"
pTitle4 = "info_title"
pIndex4 = "num"

page = currpage
%>
<td width=1 bgcolor="#990000"></td>
<td width=150 style="padding:10px" dir=rtl height=400 bgcolor="#ffffff" valign=top align=right>
<%
spics = request("spics")
if page ="pics" and spics="" then
	  sql = "SELECT num, page_name FROM pages WHERE parent = 5 or parent = 6"
    rs_tb.open sql,strconn,1,3
    while not rs_tb.eof
      pics_write = pics_write&"<a href='inner.asp?page=pics&spics="&rs_tb("num")&"'>"&rs_tb("page_name")&"</a><br>"
    rs_tb.movenext
    wend
    rs_tb.close
end if

for i=1 to pages%>
<a <%if page=eval("page"&i) then%>style="font-weight:bold"<%end if%> href="inner.asp?page=<%=eval("page"&i)%>" class=text_red><%=eval("pName"&i)%></a><br>
<%
  if page=eval("page"&i) and page<>"" then
    sql = "SELECT * FROM "&page

	  rs_tb.open sql,strconn,1,3
	  while not rs_tb.eof 
	  if int(rs_tb(eval("pIndex"&i))) = int(prev_num) then
		  isbold = "style='font-weight:bold'"
	  else
		  isbold = ""
	  end if
	  pTitle = rs_tb(eval("pTitle"&i))
	  if len(pTitle)>20 then
	  pTitle = left(rs_tb(eval("pTitle"&i)),20)&"..."
    end if
    path = page
    filepath = ""
        if page = "pics" then
          response.write "<img width='35' src='../cut/gallery/small/"&rs_tb("pic_name")&"' align='middle'>"
        end if
        response.write "<a "&isbold&" href='edit.asp?page="&page&"&num="&rs_tb(eval("pIndex"&i))&"&spics="&spics&"'><nobr>"&pTitle&"</nobr></a><br>"
		  rs_tb.movenext
	  wend
	  rs_tb.close
  end if
    if eval("page"&i) ="pics" and spics="" then
        response.write pics_write
    end if
%><br>
<%next%>
</td><td width=1 bgcolor="#990000"></td>
