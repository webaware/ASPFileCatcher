# ASPFileCatcher

This script helps Classic ASP scripts handle form posted upload files without the need for third-party components.

## Features

* uses only VBScript and standard scripting components (ADO, Scripting.* objects)
* permits access to other fields posted with form
* supports multi-value fields (e.g. checkboxes)
* files are uploaded into TEMP folder, and deleted after use unless moved
* simulates SA-FileUp so can be used as a drop-in, e.g. for testing / prototyping

## Requirements

* ADO 2.5 or greater, due to use of the ADO Stream object
* ADO constants defined, best done by loading the typelib in globals.asa

## Limitations

* not intended for use with very large (>10MB) files! use [SA-FileUp](http://safileup.com/) instead for large uploads
* NB: can't mix ASPFileCatcher and Response.Form - Response.BinaryRead restriction

## Usage

Use ASPFileCatcher as a replacement for Request.Form, to get access to all form inputs including attached files. The two are mutually exclusive; you cannot combine access to the form data through Request.Form and ASPFileCatcher, because both need to read the binary data from the form post and *there can be only one* that does that.

Include ASPFileCatcher into your script either by virtual or file include, like so:

        <!-- #include file="aspfilecatcher.asp" -->

Tell your form post to use MIME multipart form data:

        <form action="index.asp" method="post" enctype="multipart/form-data">

Use the "file" input type to attach a file to your form post:

        <input type="file" name="File1" size="40" />

Create an ASPFileCatcher object to handle form post:

        Set catcher = new ASPFileCatcher

Process each file posted, either by moving to their new destination or by reading:

        For Each f In catcher.Files
            Set fp = catcher.File(f)
            Response.Write "<p>" & Server.HTMLEncode(f) & ": " & Server.HTMLEncode(fp.FileName) & "</p>" & vbCrLf
            fp.MoveTempToPath Server.MapPath("files/" & fp.FileName)
        Next

Access non-file fields in the same way as you would normally from Request.Form, but use ASPFileCatcher instead:

        some_data = catcher.Field("some_data")

See test/index.asp for a fully worked example.
