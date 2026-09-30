
<!-- #include file="secure.inc" -->
<%
set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")
set r2=Server.CreateObject("ADODB.Recordset")
set r3=Server.CreateObject("ADODB.Recordset")
set r4=Server.CreateObject("ADODB.Recordset")
set r5=Server.CreateObject("ADODB.Recordset")

sql = "select * from [imknasim].[dbo].[conference] where theIndex="&request("item")&""
r.open sql,strconn,1,3
if r("description")&"a"<>"" then
	theDescription=replace(r("description"),"""","''")
else
	theDescription=""
end if
first=true
if not r.eof then
	theName = replace(r("theName"),"'","''")
	s = "{""theNumber"":""" & r("theIndex") & """,""theName"":""" & r("theName") & """,""banner"":""" & r("banner") & """,""logo"":""" & r("logo") & """,""isourlogo"":""" & r("isourlogo") & """,""startdate"":""" & r("startdate") & _
    """,""enddate"":""" & r("enddate") & """,""location"":""" & r("location") & """,""bkcolor"":""" & r("bkColor") & """,""textcolor"":""" & r("textColor") & _
    """,""sessionColor"":""" & r("sessionColor") & """,""splitColor"":""" & r("splitColor") & """,""hoursLeft"":""" & r("hoursLeft") & """,""textColor2"":""" & r("textColor2") &_ 
	""",""bkButton"":""" & r("bkButton") & """,""logoLight"":""" & r("logoLight") & """,""roundedCorners"":""" & r("roundedcorners") & """,""hoursframe"":""" & r("hoursframe") & """,""days"":["
	sql3 = "select distinct theday from session where conference="&request("item")&" order by theDay"
	r3.open sql3,strconn,1,3
	firstItem3=true
	while not r3.eof
		if firstItem3=false then s = s&","
		s = s&"{""theNumber"":"""&r3("theDay")&""",""splits"":["
		s = s&"{""theNumber"":0,""sessions"":["
		sql1 = "select * from session where split=0 and conference="&request("item")&" order by theOrder"
		firstItem=true
		r1.open sql1,strconn,1,3
		while not r1.eof
			if firstItem=false then s = s&","
			s = s&"{""theNumber"":"""&r1("theIndex")&""",""theName"":"""&r1("theName")&""",""showInAgenda"":"""&r1("showinagenda")&""",""items"":["
			sql2 = "select * from item where session="&r1("theIndex")&" order by theOrder"
			r2.open sql2,strconn,1,3
			firstItem1=true
			while not r2.eof
				theSponsor=""
				sql4 = "select * from sponsor where theindex="&r2("sponsership")
				r4.open sql4,strconn,1,3
				if not r4.eof then theSponsor = r4("theName")
				r4.close
				if firstItem1=false then s = s&","
				s = s&"{""theNumber"":"""&r2("theIndex")&""",""theName"":"""&r2("theName")&""",""theSponsor"":"""&theSponsor&""",""fromHour"":"""&r2("fromHour")&""",""toHour"":"""&r2("toHour")&""",""lecturers"":["
				sql4 = "select * from lecturer where theItem="&r2("theIndex")
				r4.open sql4,strconn,1,3
				if not r4.eof then
					theTite=""
					theOrg=""
					sql5 = "select * from [imknasim].[dbo].[title] where theIndex="&r4("theTitle")
					r5.open sql5,strconn,1,4
					if not r5.eof then theTitle=r5("theName")
					r5.close
					sql5 = "select * from organization where theIndex="&r4("theorganization")
					r5.open sql5,strconn,1,4
					if not r5.eof then theOrg=r5("theName")
					r5.close
					s = s & "{""theName"":"""&r4("theName")&""",""theJob"":"""&r4("theJob")&""",""theTitle"":"""&theTitle&""",""theOrg"":"""&theOrg&"""}"
				end if
				r4.close
				s =s&"]}"
				firstItem1=false
				r2.movenext
			wend
			s=s&"]}"
			r2.close
			firstItem=false
			r1.movenext
		wend
		s=s&"]}]}"
		r1.close
		firstItem3 = false
	r3.movenext
	wend
	r3.close
	s = s & "],""sponsors"":["
	sql3 = "select * from consponser where conference="&request("item")
	r3.open sql3,strconn,1,3
	firstItem="True"
	while not r3.eof
		if firstItem="False" then s = s & ","
		theSponser = "לא ידוע"
		theImage = ""
		sql4 = "select * from sponsor where theIndex="&r3("sponsor")
		r4.open sql4,strconn,1,3
		if not r4.eof then 
			theSponsor = r4("theName")
			theImage = r4("theImage")
		end if
		r4.close
		s = s & "{""theName"":"""&theSponsor&""",""theImage"":"""&theImage&"""}"
		r3.movenext
		firstItem="False"
	wend
	r3.close
	s = s & "]}"
else
	s = "{}"
end if
r.close
response.write s
%>


