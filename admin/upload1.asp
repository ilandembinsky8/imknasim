
<%

	'response.write "aaa"
	'response.end

	Set Upload = Server.CreateObject("Persits.Upload.1")

	Upload.OverwriteFiles = false
	On Error Resume Next
	Server.ScriptTimeout=20000
	Count = Upload.Save
%>
<HTML>
<head>
    <meta charset="utf-8" />
</head>
<BODY BGCOLOR="#FFFFFF">
<% If Err <> 0 Then %>

	<FONT SIZE=3 FACE="Arial" COLOR=#0020A0>
	<H3>The following error occured while uploading:</h3>
	</FONT>

	<FONT SIZE=3 FACE="Arial" COLOR=#FF2020>
	<h2>"<% = Err.Description %>"</h2>
	</FONT>

	<FONT SIZE=2 FACE="Arial" COLOR="#0020A0">
	Please <A HREF="upload1.asp">try again</A>.
	</FONT>

<% Else
filepath1 = Upload.Form("filepath1")
filepath2 = Upload.Form("filepath2")
fieldname = Upload.Form("fieldname")
thefolder = upload.Form("thefolder")
'response.Write fieldname
'response.end
For Each File in Upload.Files

fullFile = File.fileName
x = len(File.fileName)

y=" "
z = x
while y<>"." and x>0
    y = mid(file.fileName,x,1)
    x = x - 1
wend
w = z - x - 1

extension = right(fullFile,w)
fullFile = left(fullFile,len(fullFile)-4)

Dim max,min
max=9
min=0
Randomize
s = ""
for i=1 to 6
    s = s &Int((max-min+1)*Rnd+min)
next

File.SaveAs "d:\yan\berger_sisters\imknasim\images\" &s&"."&extension
RegFile = File.filename
bigfile = File.filename
'response.Write BigFile
'response.end
next

%>
<div style="float:right;width:100%">
    <div style="margin:0px auto;width:440px;height:200px;background-color:#DADADA;box-sizing:border-box;padding:20px;text-align:center">
    <div style="float:right;width:100%;font-size:24px">הקובץ עלה בהצלחה</div>
    <div style="float:right;width:200px;margin-right:120px"><input style="width:200px;height:40px;background-color:#EAEAEA;box-sizing:border-box;border-radius:8px;border:0px;margin-top:20px;font-size:24px"type=button value="סגור חלון" onclick="window.opener.document.forms[0].<%=fieldname%>.value='<%=RegFile%>';window.close()"></div>
</div>
<% End If %>


</BODY>
</HTML>