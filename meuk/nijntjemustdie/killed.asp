<%set conn = nothing%>
<html>
<HEAD>
	<TITLE> NIJNTJE MUST DIE / NIJNTJE MOET DOOD </TITLE>
	<SCRIPT language=JavaScript>
	window.defaultStatus = "nijntje must die";
	</SCRIPT>
	<style type="text/css">
	<!--
	BODY{
		background-color: #ff4000;
		font-family: verdana;
		font-size: 12px;
		color: #FF4000;
		margin: 0;
		margin-top: 20;
		scrollbar-track-color: #cccccc;
		scrollbar-shadow-color: #cccccc;
		scrollbar-highlight-color: #cccccc;
		scrollbar-3dlight-color: #cccccc;
		scrollbar-base-color:  #cccccc;
		scrollbar-arrow-color: #000000;
		scrollbar-darkshadow-color: #cccccc;	
		scrollbar-face-color: #999999;
		}
	
	TABLE	{
		font-family: verdana;
		font-size: 12;
		color: #FF4000; 
		}
	
	.tablebg {
		background-color: #ff4000;
		}
		
	img	{
		border-width: 0px;
		}
	
	A	{
		color: #000000;
		font-size: 14px;
		text-decoration: underline; 
		}
	
	A:HOVER	{
		color : #892739;
		text-decoration: blink;
		}
	
	H1 {
		color: #FF4000;
		font-size: 12px;
		}

	-->
	</style>
</HEAD>
<BODY>
<%
set fs = Server.CreateObject("Scripting.FileSystemObject")
set strFolder = fs.GetFolder(Server.MapPath("."))
set strFiles = strFolder.Files
blaataantal = strFiles.Count

dim blaat, blaataantal, blaat2

blaat = request.querystring("welkmafplaatjezalikdezekeereenslatenzienaandiehansworstendiehierkomenkijken")
if blaat = "" then
	blaat = blaataantal
end if

if blaat = 1 then
	blaat2 = blaataantal
else
	blaat2 = blaat-1
end if

set strFiles = nothing
set strFolder = nothing
set fs = nothing
%>
<table align="center" width="640" height="522">
	<tr>
		<td align="center" valign="top" background="background.gif" class="tablebg">
		
		<table cellspacing="0" cellpadding="0" border="0" width="583">
			<tr>
			    <td colspan="2"><img src="http://www.3l.nl/inc/img/spacer.gif" width="200" height="22"></td>
			</tr>
			<tr>
			    <td valign="top" width="260" height="370"><H1>Nijntje moet dood. Waarom?</H1>
		        <p>Omdat nijntje een mij iets te vrolijk <br>stom wit kut konijn is</p></td>
			    <td valign="top" width="323" height="370">
					<img src="<%=blaat%>.jpg" alt="Nijntje must die - nummertje <%=blaat%>">
			  </td>
			</tr>
		</table>
		
		<br><br><br>
		
		<table cellspacing="0" cellpadding="0" border="0" width="583">
			<tr>
				<td align="right">
					<a href="killed.asp?welkmafplaatjezalikdezekeereenslatenzienaandiehansworstendiehierkomenkijken=<%=blaat2%>">>></a>
					<br>
			  </td>
			</tr>
		</table>
		
		</td>
	</tr>
</table>
</BODY>
</HTML>