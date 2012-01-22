<%@ language="vbscript" codepage=65001 %>
<%
Option Explicit
Response.CharSet = "utf-8"
%>
<!-- #include file="../aspfilecatcher.asp" -->
<%
Class ThisPage
	Private m_catcher
	Private m_errmsg

	'----------------------------------------------------------------------
	Private Sub Class_Initialize()
		Set m_catcher = new ASPFileCatcher
	End Sub

	'----------------------------------------------------------------------
	Public Sub display()
		If Len(m_errmsg) > 0 Then
			Response.Write "<p class='errmsg'>" & m_errmsg & "</p>" & vbCrLf
		End If

		If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
			Dim f, fdata, farraydata

			' show each non-file field's values
			For Each f In m_catcher.Fields
				fdata = m_catcher.Field(f)
				If IsArray(fdata) Then
					' field is an array (e.g. checkboxes)
					Response.Write "<p>" & Server.HTMLEncode(f) & ":"
					For Each farraydata in fdata
						Response.Write vbCrLf & "<br/>&nbsp;--&gt;&nbsp;" & Server.HTMLEncode(farraydata)
					Next
					Response.Write "</p>" & vbCrLf
				Else
					' field is scalar (e.g. text field, drop-down selection, radio button)
					Response.Write "<p>" & Server.HTMLEncode(f) & ": " & Server.HTMLEncode(fdata) & "</p>" & vbCrLf
				End If
			Next

			' show details of file fields
			For Each f In m_catcher.Files
				Dim fp
				Set fp = m_catcher.File(f)
				Response.Write "<p>" & Server.HTMLEncode(f) & ": " & Server.HTMLEncode(fp.FileName) & "</p>" & vbCrLf
				fp.MoveTempToPath Server.MapPath("files/" & fp.FileName)
			Next
		End If

		showForm
	End Sub

	'----------------------------------------------------------------------
	Private Sub showForm()
	%>
		<form action="index.asp" method="post" enctype="multipart/form-data">
		<table class='inputform'>
		<tr>
			<th>Your Remark:</th>
			<td><input tabindex="10" type="text" name="Remark" size="40" /></td>
		</tr>
		<tr>
			<th>Some Number:</th>
			<td><input tabindex="15" type="text" name="SomeNumber" size="10" /></td>
		</tr>
		<tr>
			<th>File 1:</th>
			<td><input tabindex="20" type="file" name="File1" size="40" /></td>
		</tr>
		<tr>
			<th>File 2:</th>
			<td><input tabindex="25" type="file" name="File2" size="40" /></td>
		</tr>
		<tr>
			<th>Capital of Peru:</th>
			<td><input tabindex="30" type="text" name="PeruCapital" size="10" /></td>
		</tr>
		<tr>
			<th>Multi-Choice:</th>
			<td>
				<label><input tabindex="40" type="checkbox" name="MultiChoice" value="red" />&nbsp;red</label>&nbsp;&nbsp;
				<label><input tabindex="40" type="checkbox" name="MultiChoice" value="yellow" />&nbsp;yellow</label>&nbsp;&nbsp;
				<label><input tabindex="40" type="checkbox" name="MultiChoice" value="pink" />&nbsp;pink</label>&nbsp;&nbsp;
				<label><input tabindex="40" type="checkbox" name="MultiChoice" value="green" />&nbsp;green</label>&nbsp;&nbsp;
				<label><input tabindex="40" type="checkbox" name="MultiChoice" value="purple" />&nbsp;purple</label>&nbsp;&nbsp;
				<label><input tabindex="40" type="checkbox" name="MultiChoice" value="orange" />&nbsp;orange</label>&nbsp;&nbsp;
				<label><input tabindex="40" type="checkbox" name="MultiChoice" value="blue" />&nbsp;blue</label>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td><input tabindex="99" type="submit" value="Send" /></td>
		</tr>
		</table>
		</form>
	<%
	End Sub
End Class

Dim p : Set p = new ThisPage
%>
<!DOCTYPE html>
<html lang="en-au">

<head>
<title>Test File Uploads</title>
<meta charset="utf-8" />
<link rel="stylesheet" href="simple.css" />
</head>

<body>
<% p.display %>
</body>
</html>
