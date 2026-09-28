
<%
	Set Upload = Server.CreateObject("Persits.Upload.1")
	Upload.OverwriteFiles = False
	On Error Resume Next

	'Upload.SetMaxSize 1048576	 Limit files to 1MB
	Upload.RegisterServer "C:\Program Files\Persits Software\AspUpload\Bin\AspUpload.dll"
	Upload.SaveVirtual "/images"
	'response.Write Server.MapPath("/")
%>
<HTML>
<BODY BGCOLOR="#FFFFFF">
<CENTER>

<% If Err <> 0 Then %>

	<FONT SIZE=3 FACE="Arial" COLOR=#0020A0>
	<H3>The following error occured while uploading:</h3>
	</FONT>

	<FONT SIZE=3 FACE="Arial" COLOR=#FF2020>
	<h2>"<% = Err.Description %>"</h2>
	</FONT>

	<FONT SIZE=2 FACE="Arial" COLOR="#0020A0">
	Please <A HREF="demo1.asp">try again</A>.
	</FONT>

<% Else %>
<FONT SIZE=3 FACE="Arial" COLOR=#0020A0>
<h2>Success! <% = Count %> file(s) have been uploaded.</h2>
</FONT>

<FONT SIZE=3 FACE="Arial" COLOR=#0020A0>
<TABLE BORDER=1 CELLPADDING=3 CELLSPACING=0>
<TH BGCOLOR="#FFFF00">Uploaded File</TH><TH BGCOLOR="#FFFF00">Size</TH><TH BGCOLOR="#FFFF00">Original Size</TH><TR>
<%' For Each File in Upload.Files 
	File.SaveAs "images\try.jpg"
	Set File = Upload.Files(1)

%>
		<TD ALIGN=CENTER>
			<IMG SRC="<% = File.Folder%><% = File.FileName%>"><BR><B><% = File.Folder%><% = File.FileName%></B><BR>
			(<% = File.ImageWidth %> x <% = File.ImageHeight %> pixels)
		
		</TD>
	<TD ALIGN=RIGHT VALIGN="TOP"><% =File.Size %> bytes</TD>
	<TD ALIGN=RIGHT VALIGN="TOP"><% =File.OriginalSize %> bytes</TD><TR>
<%' Next %>
</TABLE>
</FONT>
<P>
<FONT SIZE=2 FACE="Arial" COLOR=#0020A0>
Click <A HREF="demo1.asp">here</A> to upload more files.
</FONT>
<%
 End If %>
</CENTER>
</BODY>
</HTML>