<%
' aspfilecatcher.asp version 1.0.4 2012-01-22
'
' handle form posted upload files in ASP without third-party components
'
' * requires ADO 2.5 or greater, due to use of the ADO Stream object
' * not intended for use with very large (>10MB) files!
' * files are uploaded into TEMP folder, and deleted after use unless moved
' * supports multi-value fields (e.g. checkboxes)
' * NB: can't mix ASPFileCatcher and Response.Form - Response.BinaryRead restriction
'
' copyright (c) 2008-2012 WebAware Pty Ltd
' https://github.com/webaware/ASPFileCatcher
'----------------------------------------------------------------------
' This library is free software; you can redistribute it and/or
' modify it under the terms of the GNU Lesser General Public
' License as published by the Free Software Foundation; either
' version 2.1 of the License, or (at your option) any later version.
'
' This library is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' Lesser General Public License for more details.
'
' You should have received a copy of the GNU Lesser General Public
' License along with this library; if not, write to the Free Software
' Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
'
' Full license: http://www.webaware.com.au/free/license.htm
'----------------------------------------------------------------------
' ASPFileCatcher Properties:
' Fields				get an array of field names that aren't files
' Field					get data (string) sent for named field
' Files					get an array of file field names
' File					get ASPFileCatcherFile object for named file field
'----------------------------------------------------------------------
' ASPFileCatcherFile Properties:
' FieldName				name of file field
' FileName				filename from source system
' TempFilePath			full path to uploaded file in temporary folder
' ContentType			MIME content type as sent by browser
'----------------------------------------------------------------------
' ASPFileCatcherFile Methods:
' MoveTempToPath		move and rename file from temporary folder to new full path
' DeleteTempFile		delete temporary file if it hasn't been moved
'----------------------------------------------------------------------

Class ASPFileCatcher
	Private m_delimiterB
	Private m_files
	Private m_fields
	Private m_postdata
	Private m_postdata_pos
	Private m_postdata_len

	' cache objects for simulating SA-FileUp component (for code compatibility)
	Private m_form

	'----------------------------------------------------------------------
	Private Sub Class_Initialize()
		' create a dictionary object for storing the file objects
		Set m_files = Server.CreateObject("Scripting.Dictionary")

		' create a dictionary object for storing the non-file field data
		Set m_fields = Server.CreateObject("Scripting.Dictionary")

		If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
			' read the whole form post into a variant
			m_postdata = Request.BinaryRead(Request.TotalBytes)
			m_postdata_pos = 1
			m_postdata_len = Request.TotalBytes

			' grab the field delimiter
			m_delimiterB = readResponseLineB

			' process fields until delimiter followed by "--"
			Do
				processField
			Loop Until foundEndDelimiter

			' release resources
			m_postdata = Empty
			m_postdata_len = 0
		End If
	End Sub

	'----------------------------------------------------------------------
	Private Sub Class_Terminate()
		' if any files haven't been moved or deleted from temporary location, clean them up now
		Dim f, fp

		For Each f In m_files.Keys
			Set fp = m_files.Item(f)
			fp.DeleteTempFile
		Next
	End Sub

	'----------------------------------------------------------------------
	Public Property Get Fields()
		Fields = m_fields.Keys
	End Property

	'----------------------------------------------------------------------
	Public Property Get Field(ByVal key)
		Field = m_fields.Item(key)
	End Property

	'----------------------------------------------------------------------
	' simulate SA-FileUp component (for code compatibility)
	Public Property Get Form()
		If Not IsObject(m_form) Then
			Dim i
			Set m_form = Server.CreateObject("Scripting.Dictionary")

			For Each i In m_fields
				m_form(i) = m_fields.Item(i)
			Next

			For Each i In m_files
				Set m_form(i) = m_files.Item(i)
			Next
		End If

		Set Form = m_form
	End Property

	'----------------------------------------------------------------------
	Public Property Get Files()
		Files = m_files.Keys
	End Property

	'----------------------------------------------------------------------
	Public Property Get File(ByVal key)
		Set File = m_files.Item(key)
	End Property

	'----------------------------------------------------------------------
	Private Sub processField()
		Dim contentDisposition, contentType, fieldname, filename, tmpname, fileObj
		Dim re, reMatch, reMatches
		Dim delimPos, buffer, bin_data

		Set re = new RegExp

		contentDisposition = readResponseLine

		re.Pattern = "name=""([^""]*)"""
		Set reMatches = re.Execute(contentDisposition)
		fieldname = reMatches(0).SubMatches(0)

		re.Pattern = "filename=""([^""]*)"""
		Set reMatches = re.Execute(contentDisposition)
		If reMatches.count = 0 Then
			' not a file field, grab the data and move on
			retrieveFormField fieldname
			Exit Sub
		End If
		filename = reMatches(0).SubMatches(0)

		contentType = readResponseLine
		re.Pattern = "Content-Type: (.*)"
		Set reMatches = re.Execute(contentType)
		contentType = reMatches(0).SubMatches(0)

		readResponseLineB	' read line space between headers and file data

		delimPos = InStrB(m_postdata_pos, m_postdata, m_delimiterB)
		If delimPos >= m_postdata_pos Then
			If Len(filename) > 0 Then
				' convert VBScript array chunk into binary data
				Set bin_data = Server.CreateObject("ADODB.Stream")
				bin_data.Type = adTypeText
				bin_data.Open
				bin_data.WriteText MidB(m_postdata, m_postdata_pos, delimPos - m_postdata_pos - 2)
				bin_data.Position = 0
				bin_data.Type = adTypeBinary
				bin_data.Position = 2           ' Important: skip first two bytes that mark it as Unicode string!

				' open a binary stream
				Set buffer = Server.CreateObject("ADODB.Stream")
				buffer.Type = adTypeBinary
				buffer.Open

				' now write binary data to the output stream
				buffer.Write bin_data.Read
				bin_data.Close
				Set bin_data = Nothing

				' and write the output stream to a file
				tmpname = getTempName
				buffer.SaveToFile tmpname, adSaveCreateOverWrite
				buffer.Close
				Set buffer = Nothing
			End If

			' advance the data pointer
			m_postdata_pos = delimPos + LenB(m_delimiterB)

			' record in files collection
			Set fileObj = new ASPFileCatcherFile
			fileObj.FieldName = fieldname
			fileObj.FileName = filename
			fileObj.TempFilePath = tmpname
			fileObj.ContentType = contentType
			m_files.Add fieldname, fileObj
		End If
	End Sub

	'----------------------------------------------------------------------
	' get the data for a non-file field
	Private Sub retrieveFormField(fieldname)
		Dim data, delimPos

		readResponseLineB	' read line space between headers and field data

		data = readResponseLine
		If m_fields.Exists(fieldname) Then
			' need to convert scalar value into an array, so can hold multiple values for one field name
			Dim oldData, newData
			oldData = m_fields.Item(fieldname)
			If IsArray(oldData) Then
				Redim Preserve oldData(UBound(oldData) + 1)
				oldData(UBound(oldData)) = data
			Else
				oldData = Array(oldData, data)
			End If
			m_fields.Item(fieldname) = oldData
		Else
			' just store scalar value for this field name
			m_fields.Add fieldname, data
		End If

		' skip to end of delimiter
		delimPos = InStrB(m_postdata_pos, m_postdata, m_delimiterB)
		If delimPos >= m_postdata_pos Then
			m_postdata_pos = delimPos + LenB(m_delimiterB)
		Else
			m_postdata_pos = m_postdata_len
		End If
	End Sub

	'----------------------------------------------------------------------
	' read bytes from the Response up to the end of a line (CRLF)
	Private Function readResponseLine()
		Dim s

		s = readResponseLineB()
		If LenB(s) > 0 Then
			readResponseLine = strBtoU(s)
		End If
	End Function

	'----------------------------------------------------------------------
	' read bytes from the Response up to the end of a line (CRLF) as single-byte char string
	Private Function readResponseLineB()
		Dim endpos

		endpos = InstrB(m_postdata_pos, m_postdata, ChrB(&h0a))
		If endpos >= m_postdata_pos Then
			If endpos >= m_postdata_pos + 2 Then
				readResponseLineB = MidB(m_postdata, m_postdata_pos, endpos - m_postdata_pos - 1)
			End If
			m_postdata_pos = endpos + 1
		End If
	End Function

	'----------------------------------------------------------------------
	' read two bytes from the Response, and see whether it's the end of post delimiter
	Private Function foundEndDelimiter()
		Dim next2

		next2 = strBtoU(MidB(m_postdata, m_postdata_pos, 2))
		m_postdata_pos = m_postdata_pos + 2

		If next2 = "--" Then
			foundEndDelimiter = True
		Else
			foundEndDelimiter = False
		End If
	End Function

	'----------------------------------------------------------------------
	' convert a single-byte string to a Unicode string (as used in VBScript)
	Private Function strBtoU(strB)
		Dim i

		strBtoU = ""
		For i = 1 to LenB(strB)
			strBtoU = strBtoU & Chr(AscB(MidB(strB, i, 1)))
		Next
	End Function

	'----------------------------------------------------------------------
	' get a unique temporary file name (with full path)
	Private Function getTempName()
		Dim fso, tempfolder

		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		Set tempfolder = fso.GetSpecialFolder(2)	' 2 = TemporaryFolder
		getTempName = tempfolder & "\" & fso.GetTempName
	End Function
End Class

'----------------------------------------------------------------------
Class ASPFileCatcherFile
	Public FieldName
	Public FileName
	Public TempFilePath
	Public ContentType

	'----------------------------------------------------------------------
	' simulate SA-FileUp component (for code compatibility)
	Public Property Get ServerName()
		ServerName = TempFilePath
	End Property

	'----------------------------------------------------------------------
	' simulate SA-FileUp component (for code compatibility)
	Public Property Get ShortFilename()
		ShortFilename = FileName
	End Property

	'----------------------------------------------------------------------
	' simulate SA-FileUp component (for code compatibility)
	Public Property Get IsEmpty()
		IsEmpty = False
	End Property

	'----------------------------------------------------------------------
	' if the temporary file hasn't been moved yet, move it to new full path
	Public Function MoveTempToPath(newFilePath)
		Dim fso

		If Len(TempFilePath) > 0 Then
			Set fso = Server.CreateObject("Scripting.FileSystemObject")

			If fso.FileExists(newFilePath) Then
				fso.DeleteFile newFilePath
			End If
			fso.MoveFile TempFilePath, newFilePath

			TempFilePath = ""
			MoveTempToPath = True
		Else
			MoveTempToPath = False
		End If
	End Function

	'----------------------------------------------------------------------
	' if the temporary file hasn't been moved, delete it
	Public Sub DeleteTempFile()
		Dim fso
		On Error Resume Next

		If Len(TempFilePath) > 0 Then
			Set fso = Server.CreateObject("Scripting.FileSystemObject")
			If fso.FileExists(TempFilePath) Then
				fso.DeleteFile TempFilePath
			End If
		End If
	End Sub
End Class
%>
