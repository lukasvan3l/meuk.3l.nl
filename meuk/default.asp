<%@LANGUAGE = VBScript%>
<% 
Response.Buffer = TRUE
Response.Expires = 0
session.lcid = 1043
session.timeout = 30 '30 minuten
Server.ScriptTimeout = 3600 '= 60 minuten%>
<html>
<HEAD>
	<%if session("username") <> "" then%>
		<TITLE>Downloads - <%=session("username")%></TITLE>
	<%else%>
		<TITLE>Downloads - niet ingelogd</TITLE>
	<%end if%>
	<link rel="STYLESHEET" type="text/css" href="http://www.3l.nl/inc/css/zwart.css">
	<style>
	a {
		padding-left:15px;
	}
	</style>
	<script language="JavaScript" type="text/javascript">
	function confirmDelete(bestand){
		if (bestand == "")
			var agree=confirm("Weet u zeker dat dit item wilt verwijderen?");
		else
			var agree=confirm(bestand + " verwijderen?");
		if (agree)
			return true ;
		else
			return false ;
	}
	</script>
</HEAD>
<BODY>
<!-- #include file="../inc/functions.asp" -->
<%
nivo = "user"					'Dit niveau moet je zijn om te uploaden, "" betekent iedereen
nivoverwijderen = "administrator"'Dit niveau moet je zijn om te deleten, "" betekent iedereen
nivobekijken = ""				'Dit niveau moet je zijn om te kijken, "" betekent iedereen
aantalfiles = 1					'aantal bestanden dat je kunt uploaden:
maxMB = 10						'aantal mb dat een bestand mag zijn
minvrij = 20					'minimum aantal MB vrij op de schijf
fotostonen = true				'foto's tonen onderaan de pagina

if nivobekijken = "" or checkIngelogd(nivobekijken) = true then

set fs=Server.CreateObject("Scripting.FileSystemObject")
	vrijeruimte = (150-(fs.GetFolder("D:\www\3l.nl\").Size / 1024 / 1024))
set fs=nothing

if not request.querystring("d") = "" and (nivoverwijderen = "" or checkIngelogd(nivoverwijderen)=true) then
	deleteFile(replace(left(request.servervariables("SCRIPT_NAME"),instrrev(request.servervariables("SCRIPT_NAME"),"/")) & request.querystring("d"),"/","\"))
	response.redirect("default.asp?v="&request.querystring("d"))
elseif checkIngelogd(nivo) = true or nivo = "" then
	if request.querystring("a") = "upload" and vrijeruimte > minvrij then
		Set Upload = Server.CreateObject("Persits.Upload")
		Upload.ProgressID = Request.QueryString("PID")
		Upload.OverwriteFiles = False
		Upload.SetMaxSize maxMB*1000000, true
		Count = Upload.Save(Server.MapPath("."))
		response.redirect("default.asp?u="&count)
	elseif aantalfiles > 0 and vrijeruimte > minvrij then
		Dim UploadProgress, PID, barref
		Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
		PID = "PID=" & UploadProgress.CreateProgressID()
		barref = "http://www.3l.nl/admin/framebar.asp?to=10&" & PID
		%>
		<SCRIPT LANGUAGE="JavaScript">
		function ShowProgress()
		{
			strAppVersion = navigator.appVersion;
			if (document.uploadForm.Path.value != "")
			{
				if (strAppVersion.indexOf('MSIE') != -1 && strAppVersion.substr(strAppVersion.indexOf('MSIE')+5,1) > 4)
				{
					winstyle = "dialogWidth=375px; dialogHeight:130px; center:yes";
					window.showModelessDialog('<% = barref %>&b=IE',null,winstyle);
				}
				else
				{
				window.open('<% = barref %>&b=NN','','width=370,height=115', true);
				}
			}
			return true;
		}
		</SCRIPT> 
		<h1>Bestand toevoegen</h1>
		<form action="default.asp?a=upload&<% = PID %>" method="post" enctype="multipart/form-data" name="uploadForm" id="uploadForm" onsubmit="return ShowProgress();" style="margin-left:15px;">
		<%
		if not request.querystring("u") = "" then
			response.write("<p align=""center"" style=""color:red;"">Er zijn "&request.querystring("u")&" bestand(en) geupload.</p>")
		elseif not request.querystring("v") = "" then
			response.write("<p align=""center"" style=""color:red;"">Bestand "&request.querystring("v")&" is verwijderd.</p>")
		else
			response.write("<p>Er is nog "&formatnumber(vrijeruimte - minvrij,0)&" Mb vrij.</p>")
		end if
		%>
		<input type="file" name="Path" size="30"><br>
		<%for i = 2 to aantalfiles%>
			<input type="file" name="Path<%=i%>" size="30"><br>
		<%next%>
		<input type="submit" value="Upload">
		</form>
		<%
	end if
end if




writePics = ""
writeFolders = ""
writeFiles = ""

set fs=Server.CreateObject("Scripting.FileSystemObject")
set fo=fs.GetFolder(server.mappath(request.servervariables("SCRIPT_NAME")&"/.."))

for each x in fo.SubFolders
	writeFolders = writeFolders & "<p><a href=""" & x.Name & """>" & x.Name & "</a></p>"&vbcrlf
next


for each x in fo.files
	if lcase(right(x,4)) = ".jpg" or lcase(right(x,4)) = ".gif" then
		on error resume next
		Set jpeg = Server.CreateObject("Persits.Jpeg")
		jpeg.Open( x )
		if jpeg.OriginalWidth > 800 then
			width = "800"
		else
			width = jpeg.OriginalWidth
		end if
		set jpeg = nothing
		writePics = writePics & "<p align=""center""><a href="""&x.name&""">"
		writePics = writePics & "<img src="""&x.name&""" alt="""&x.name&""" width="""&width&"""></a></p>"&vbcrlf
	end if
	if not lcase(x.name)="default.asp" then
		writeFiles = writeFiles & "<p>"
		if checkIngelogd(nivoverwijderen) = true or nivoverwijderen = "" then
			writeFiles = writeFiles & "<a href=""default.asp?d="&x.name&""" onClick=""return confirmDelete('"&x.name&"')"">"
			writeFiles = writeFiles & "<img src=""http://www.3l.nl/admin/images/delete.gif"" alt=""Verwijder"" border=0></a>"
		end if
		writeFiles = writeFiles & "<a href="""&x.name&""">"&x.name&"</a></p>"&vbcrlf
	end if
next


if not writeFolders = "" then
	response.write("<h1>Mappen</h1>")
	response.write("<p>Klik op de naam van een subfolder om de bestanden te bekijken.</p>")
	response.write(writeFolders)
	response.write("<br /><br />")
end if

if not writeFiles = "" then
	response.write("<h1>Bestanden</h1>")
	response.write("<p>Klik op de naam van een bestand om het bestand te bekijken.<br>")
	response.write("Om een bestand op te slaan klik met je rechter muisknop en kies 'doel opslaan als...'</p>")
	response.write(writeFiles)
	response.write("<br /><br />")
end if

if not writePics = "" and fotostonen = true then
	response.write("<h1>Foto's</h1>")
	response.write("<p>Klik op de foto om het origineel te bekijken.<br>Om een bestand op te slaan klik met je rechter muisknop en kies 'doel opslaan als...'</p>")
	response.write(writePics)
end if

if writePics = "" AND writeFiles = "" AND writeFolders = "" then
	response.write("<h1>Leeg</h1><p>Er bevinden zich geen mappen of bestanden in deze directory.</p>")
end if

set fo=nothing
set fs=nothing


else
	response.write("<h1>Verboden toegang</h1>Voor deze pagina moet je als "&nivobekijken&" <a href=""http://www.3l.nl/login.asp"">inloggen</a>.")
end if
%>
</BODY>
</HTML>