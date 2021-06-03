'
'    TexMathsTools
'
'	 Copyright (C) 2012-2020 Roland Baudin (roland65@free.fr)
'    Based on the work of Geoffroy Piroux (gpiroux@gmail.com)
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA
'
'	 Various general macros


' Force variable declaration
Option Explicit


' Get TexMaths version
Function GetVersion() as String

	' Get a list of all installed extensions
	Dim oPackageInfoProvider as Variant, list as Variant
	oPackageInfoProvider = GetDefaultContext.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
	list = oPackageInfoProvider.getExtensionList()

	' Get TexMaths version from the list
	Dim version as String
	Dim i as Integer
	For i = 0 to UBound(list(), 1)

		If list(i)(0) = "org.roland65.texmaths" Then
			version = list(i)(1)
			Exit For
		End If

	Next i

	GetVersion = version

End Function


' Get LibreOffice version (i.e. major version number, ex: 6.2)
Function GetLOVersion() as String

	Dim aConfigProvider as Variant
	Dim oNode as Variant
	Dim args(0) as new com.sun.star.beans.PropertyValue

	aConfigProvider = createUnoService("com.sun.star.configuration.ConfigurationProvider")
	args(0).Name = "nodepath"
	args(0).Value = "/org.openoffice.Setup/Product"
	   
	oNode = aConfigProvider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", args())
	GetLOVersion = oNode.getbyname("ooSetupVersion")
 
End Function


' Get LibreOffice About version (i.e. complete version number, ex: 6.2.4.1)
Function GetLOAboutVersion() as String

	Dim aConfigProvider as Variant
	Dim oNode as Variant
	Dim args(0) as new com.sun.star.beans.PropertyValue

	aConfigProvider = createUnoService("com.sun.star.configuration.ConfigurationProvider")
	args(0).Name = "nodepath"
	args(0).Value = "/org.openoffice.Setup/Product"
	   
	oNode = aConfigProvider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", args())
	GetLOAboutVersion = oNode.getbyname("ooSetupVersionAboutBox")
 
End Function


' Get the script directory path depending on the system
' On Windows, for LibreOffice version < 6.2.4.1, we use a non standard place due to a bug in LibreOffice
' where it is not possible to execute a batch file when the path name contains spaces
' Since LibreOffice 6.2.4.1, we use the standard place
Function GetScriptDir() as String

	' Windows
	If getGUIType() = 1 Then

		Dim sVersion, sAboutVersion as String
		sVersion = GetLOVersion()
		sAboutVersion = Right(GetLOAboutVersion(), 3)

		' If LibreOffice version >= 6.2.4.1, standard place
	    If ( Val(sVersion) = 6.2 And  Val(sAboutVersion) >= 4.1 ) Or ( Val(sVersion) > 6.2 ) Then

			GetScriptDir = ConvertFromURL( glb_UserPath )
	    
	    ' If LibreOffice version < 6.2.4.1, non standard place
	    Else
	    	    
			GetScriptDir = Environ("HOMEDRIVE") & "\" & "texmaths-" & RemoveSpaces( Environ( "USERNAME" ) & "\" )
		
		End If
		
 	' Linux or MacOSX
 	Else
	
		GetScriptDir = glb_UserPath
	
	End If

End Function


' Get the base directory of a file given by its URL
' The return base directory is a path, not and URL
' Adapted from a function by ThierryM 
Function GetBaseDir(sFilePath as String) as String

	If sFilePath = "" Then

		GetBaseDir = ""
	
	Else
	
		Dim sURL as String	
		sURL = ConvertToURL(sFilePath)
	
		Dim pos as Integer
		pos = len(sURL)
		While mid(sURL, pos, 1) <> "/"
			pos = pos - 1
		Wend
		
		GetBaseDir = ConvertFromURL(mid(sURL, 1, pos))
	
	End If

End Function


' Get the script file path depending on the system
Function GetScriptPath() as String

	If getGUIType() = 1 Then ' Windows
		
		GetScriptPath = GetScriptDir() & "TexMaths-" & GetVersion() & ".bat"
		
	Else ' Linux or MacOSX

		GetScriptPath = GetScriptDir() & "TexMaths-" & GetVersion() & ".sh"

	End If

End Function



' Execute the Unix command 'uname' and return the result into a string
' On Mac OS X, this should return the string "Darwin"
Function GetUname() as String

	Dim sMsg as String, sFilePath as String

	' Execute the uname command
	sFilePath =  ConvertFromURL(glb_TmpPath) & "tmpuname.txt"	
	Shell( "sh -c", 2,  "'" & "uname > " & """" & sFilePath & """" & "'", TRUE )

	' File does not exist 
	If Not FileExists( sFilePath ) Then

		MsgBox( _("Error: can't find file ") & sFilePath, 0, "TexMaths")
		Exit Function

	End If
	
	' Read first line of file
	sMsg = ""
	Dim iNumber as Integer
	Dim sLine as String
	iNumber = Freefile
	Open sFilePath For Input As iNumber

		If Not EOF(iNumber) Then
  			Line Input #iNumber, sLine
			sMsg = sLine
		End If

	Close #iNumber

	' Output the result
	GetUname = sMsg

End Function


' Return TRUE if dvisvgm supports XeLaTeX
' Return FALSE if dvisvgm doesn't support XeLaTeX
Function dvisvgmSupportsXelatex() as Boolean

	Dim sMsg as String, sFilePath as String, sCmd as String
	Dim oSystemInfo as Variant
	oSystemInfo = GetConfigAccess( "/ooo.ext.texmaths.Registry/SystemInfo", TRUE)	

	' Check if dvisvgm exists
	If Not FileExists(oSystemInfo.DvisvgmPath) Then
		dvisvgmSupportsXelatex = FALSE
		Exit Function
	End If

	' File path
	sFilePath =  ConvertFromURL(glb_TmpPath) & "tmpcmd.txt"	

	' Command to execute
	sCmd =  oSystemInfo.DvisvgmPath

	' Windows
	If getGUIType() = 1 Then

		' Open service file and an output stream
		Dim cURL as String, str as String	
		Dim oFileAccess as Variant, oTextStream as Variant
		oFileAccess = createUnoService("com.sun.star.ucb.SimpleFileAccess")
		oTextStream  = createUnoService("com.sun.star.io.TextOutputStream")
		
		' Generate the script "version-XX.bat" if it doesn't exist yet
		' This script is used to know the version of a given program
		Dim sScriptName as String
		sScriptName = GetScriptDir() & "version-" & GetVersion() & ".bat"
		
		If Not FileExists(sScriptName ) Then
			cURL = ConvertToURL( sScriptName )		
			str = "@echo off"  & chr(10) &_
			chr(10) &_
			"rem This script is part of the TexMaths package" & chr(10) &_
			"rem http://roland65.free.fr/texmaths" & chr(10) &_
			"rem" & chr(10) &_
			"rem Roland Baudin (roland65@free.fr)" & chr(10) &_
			chr(10) &_
			"""" & sCmd & """" & " --version > " & """" & sFilePath & """" & chr(10)
			oTextStream.setOutputStream(oFileAccess.openFileWrite(cURL))
			oTextStream.writeString( str )
			oTextStream.closeOutput()
		End If
		 
		' Execute the command
		Shell(ConvertToURL( sScriptName ), 2, "", TRUE)

	' Linux or MacOS X
	Else
		
		' Execute the command
		Shell( "sh -c", 2,  "'" & """" & sCmd & """" & " --version > " & """" & sFilePath & """" & "'", TRUE )

	End If

	' Output file does not exist 
	If Not FileExists( sFilePath ) Then

		MsgBox( _("Error: can't find file ") & sFilePath, 0, "TexMaths")
		Exit Function

	End If
	
	' Read first line of output file
	sMsg = ""
	Dim iNumber as Integer
	Dim sLine as String
	iNumber = Freefile
	Open sFilePath For Input As iNumber

		If Not EOF(iNumber) Then
  			Line Input #iNumber, sLine
			sMsg = sLine
		End If
	Close #iNumber

	' Extract version number as "x.y.z"
	Dim version as String

	' In some versions of dvisvgm the version string is "dvisvgm (TeX Live) x.y.z"
	If len(sMsg) > 20 Then	
		version = mid(sMsg, 19)

	' In other versions it's "dvisvgm x.y.z"
	Else
		version = mid(sMsg, 8)
	End If

	' Get the major and minor version numbers
	Dim vstr() as String,  major as Integer, minor as Integer
	vstr = split(version, ".", 3)
	major = vstr(0)
	minor = vstr(1)

	' Get result
	Dim res as Boolean : res = FALSE
	If major >=2 Then
		res = TRUE
	Else
		If minor >= 16 Then
			res = TRUE
		End If
	End If

	' Return result
	dvisvgmSupportsXelatex = res

End Function


' Remove spaces in string
Function RemoveSpaces( ByVal str As String ) As String
	
	Dim result as String
	Dim c as String
	Dim i as Integer
	
	result = ""
   	For i = 1 To Len( str )
      c = Mid( str, i, 1 )
      If c <> " " Then
         result = result & c
      EndIf
   Next
   RemoveSpaces = result

End Function


' Replace newline characters (chr(10)) with "�"
Function EncodeNewline( ByVal str As String ) As String
	
	Dim result as String
	Dim c as String
	Dim i as Integer
	
	result = ""
   	For i = 1 To Len( str )	  
   	  c = Mid( str, i, 1 )
      If c = chr(10) Then
      	 c = "�"
      EndIf
	  result = result & c
   Next
   EncodeNewline = result

End Function


' Replace "�" characters with newline (chr(10))
Function DecodeNewline( ByVal str As String ) As String
	
	Dim result as String
	Dim c as String
	Dim i as Integer
	
	result = ""
   	For i = 1 To Len( str )
   	  c = Mid( str, i, 1 )
      If c = "�" Then
      	 c = chr(10)
      EndIf
	  result = result & c
   Next
   DecodeNewline = result

End Function



' Read program path from the system path
Function ReadPgmPath(pgm as String) as String

	Dim sFilePath as String, sShellCommand as String, sShellArg as String, sTmpPath as String

	' Windows
	If getGUIType() = 1 Then
   			
   		' Create the script "which-XX.bat" if it doesn't exist yet
   		' This script is used to know if a program is in the system path
   		Dim sScriptName as String
   		sScriptName = GetScriptDir() & "which-" & GetVersion() & ".bat"
   		
   		If Not FileExists( sScriptName ) Then

			Dim cURL as String, str as String
			
			' Open service file and an output stream
			Dim oFileAccess as Variant, oTextStream as Variant
			oFileAccess = createUnoService("com.sun.star.ucb.SimpleFileAccess")
			oTextStream  = createUnoService("com.sun.star.io.TextOutputStream")
		
			' Generate the "which-XX.bat" script in the appropriate directory
			cURL = ConvertToURL( sScriptName )		
		
			str = "@echo off" & chr(10) &_
			chr(10) &_
			"rem This script is part of the TexMaths package" & chr(10) &_
			"rem http://roland65.free.fr/texmaths" & chr(10) &_
			"rem" & chr(10) &_
			"rem Roland Baudin (roland65@free.fr)" & chr(10) &_
			chr(10) &_
			"rem Process the options" & chr(10) &_
			"set FILE=%~n1" & chr(10) &_
			"set TMPPATH=%2" & chr(10) &_
			chr(10) &_	
			"rem Convert TMPPATH from URL to path" & chr(10) &_
			"setlocal enabledelayedexpansion" & chr(10) &_
			"set TMPPATH=%TMPPATH:file:///=%" & chr(10) &_
			"set TMPPATH=%TMPPATH:/=\%" & chr(10) &_
			"set TMPPATH=!TMPPATH:%%20= !" & chr(10) &_
			"set TMPPATH=""%TMPPATH%""" & chr(10) &_
			chr(10) &_
			"rem Search path" & chr(10) &_
			"@for %%e in (%PATHEXT%) do @for %%i in (%FILE%%%e) do @if NOT ""%%~$PATH:i""=="""" echo %%~$PATH:i>%TMPPATH%" & chr(10)

			oTextStream.setOutputStream(oFileAccess.openFileWrite(cURL))
			oTextStream.writeString( str )
			oTextStream.closeOutput()
		 
		End If

		' Shell command
		sShellCommand = ConvertToURL( sScriptName )
				
		' Tmp file path
		sTmpPath =  ConvertToURL( GetScriptDir() & "tmppath.txt" )

		' Shell argument
		sShellArg = pgm & " " & sTmpPath

   		' Execute the command
   		Shell( sShellCommand, 2, sShellArg, TRUE )
   		
	' Linux or MacOSX
	Else
	
		' Tmp file path
		sTmpPath =  ConvertFromURL(glb_TmpPath) & "tmppath.txt"

		' Execute the which command
		Shell( "sh -c", 2,  "'" & "which " & pgm & "> " & """" & sTmpPath & """" & "'", TRUE )

		' In MacOSX, the MacTeX path is probably not set, so try the two most common paths i.e. /Library/TeX/texbin/ and /usr/local/bin/
		If GetUname() = "Darwin" And FileLen( sTmpPath ) = 0 Then
	
			Dim pgmpath as String
			pgmpath = "/Library/TeX/texbin/" & pgm
			Shell( "sh -c", 2,  "'" & "which " & pgmpath & "> " & """" & sTmpPath & """" & "'", TRUE )
			
			If FileLen( sTmpPath ) = 0 Then
				pgmpath = "/usr/local/bin/" & pgm
				Shell( "sh -c", 2,  "'" & "which " & pgmpath & "> " & """" & sTmpPath & """" & "'", TRUE )
			End If
		
		End If

	End If

	' Tmp file does not exist 
	If Not FileExists( sTmpPath ) Then

		MsgBox( _("Error: can't find file ") & sTmpPath, 0, "TexMaths")
		Exit Function

	End If
	
	' Read first line of tmp file
	sFilePath = ""
	Dim iNumber as Integer
	Dim sLine as String
	iNumber = Freefile
	Open sTmpPath For Input As iNumber

		If Not EOF(iNumber) Then
  			Line Input #iNumber, sLine
			sFilePath = sLine
		End If

	Close #iNumber

	' Return program path
	If sFilePath = "" Then
		ReadPgmPath = ""
	Else
		ReadPgmPath = sFilePath
	End If
	
End Function


' Read a text file
' The path variable must be terminated by a path separator
Function ReadTextFile(file as String, path as String) as String

	Dim sMsg as String, sFilePath as String
	 
	sFilePath = ConvertToURL( path & file )
	If Not FileExists( sFilePath ) Then

		MsgBox( _("Error: can't find file ") & file, 0, "TexMaths")
		Exit Function

	End If
	
	Dim iNumber as Integer
	Dim sLine as String
	iNumber = Freefile
	sMsg = ""
	Open sFilePath For Input As iNumber
 	While Not EOF(iNumber)

  		Line Input #iNumber, sLine
		sMsg = sMsg & sLine & chr(10)

	Wend
	Close #iNumber

	ReadTextFile = sMsg

End Function


' Read a text file encoded in UTF-8
' The path variable must be terminated by a path separator
Function ReadTextFileUtf8( file as String , path as String) as String

	Dim sMsg as String, sFilePath as String
	 
	sFilePath = ConvertToURL( path & file )
	If Not FileExists( sFilePath ) Then

		MsgBox( _("Error: can't find file ") & file, 0, "TexMaths")
		Exit Function

	End If

	Dim oTextFile as Variant, oFileAccess as Variant, oFileStream as Variant
	Dim sLine As String

	oFileAccess = createUnoService("com.sun.star.ucb.SimpleFileAccess")
	oFileStream = oFileAccess.openFileRead(sFilePath)
	oTextFile = createUnoService("com.sun.star.io.TextInputStream")
	oTextFile.InputStream = oFileStream

	sMsg = ""
	Do While Not oTextFile.IsEOF
  		sLine = oTextFile.readLine
		sMsg = sMsg & sLine & chr(10)
	Loop

	oFileStream.closeInput
	oTextFile.closeInput

	ReadTextFileUtf8 = sMsg

End Function



' Import graphic from URL into the clipboard
' Inspired from OOoForums DannyB's code
' Return TRUE if success, else return FALSE
Function ImportGraphicIntoClipboard(cURL as String, sEqFormat as String, sEqDPI as String, sEqTransp as String) as Boolean

	Dim sLine as String
	Dim str1() as String
	Dim str2 as String
	Dim X as String, Y as String, W as String, H as String 
	Dim pos1 as Integer, pos2 as Integer
	Dim cURL_ as String
	Dim iNumber as Integer
	Dim iNumber_ as Integer

	' Initialize return code to success
	ImportGraphicIntoClipboard = TRUE

	' Copy the URL of the image file
	cURL_ = cURL

	' If SVG image, then eventually add an opaque rectangle below the image
	' to improve the selection usability
	If sEqFormat = "svg" Then

		' URL of the output file
		str1 = Split(cURL, ".svg")
		cURL_ = str1(0) & "_.svg"

		' Read and write image file at the same time
		iNumber = Freefile
		iNumber_ = Freefile + 1
		sLine = ""

		If FileExists(cURL) Then
		
			' Get image transparency
			Dim opacity as String
	
			If sEqTransp = "TRUE" Then
				opacity = "0"
			Else
				opacity = "1"
			End If
	
			Open cURL For Input As iNumber
			Open cURL_ For Output As iNumber_

			Do While not Eof(iNumber)

				Line Input #iNumber, sLine		
			
				' Recopy the line
				Print #iNumber_, sLine
						
				' Add a rectangle with appropriate opacity below the equation
				If InStr(sLine, "<svg") Then
											
					If InStr(sLine, "viewBox") = 0 Or InStr(sLine, "viewBox='0 0 0 0'") <> 0 Then
					
						ErrorDialog( _("LaTeX code was successfully compiled but equation image is empty, please check the equation syntax...") )
						ImportGraphicIntoClipboard = FALSE
						Exit Function
																
					End If	
							
					' Get rectangle coordinates and size 
					pos1 = InStr(sLine, "viewBox") + 9
					pos2 = InStr(pos1, sLine, "'")			
					str2 = Mid(sLine, pos1, pos2 - pos1)
					str1 = Split(str2," ")				
					X = str1(0)
					Y = str1(1)
					W = str1(2)
					H = str1(3)		
					
					' Shrink the rectangle a bit to be sure it is invisible
					' and only write to the file if the rectangle has non zero size 
					If Val(W) > 1.0 And Val(H) > 1.0 Then
						
						' Shrinked rectangle coordinates				
						X = LTrim( Str(Val(X) + 0.5) ) ' LTrim() removes leading spaces
						Y = LTrim( Str(Val(Y) + 0.5) )
						W = LTrim( Str(Val(W) - 1.0) )
						H = LTrim( Str(Val(H) - 1.0) )
					
						str2 =  "<rect fill=""#ffffff"" x=""" & X & """ y=""" & Y & """ width=""" & W & """ height=""" & H & """ style=""fill-opacity:" & opacity &  """/>"
										
						' Write the rectangle
						Print #iNumber_, str2
					
					End If
				
				End If

			Loop

			Close #iNumber
			Close #iNumber_

		' File not found
		Else

			MsgBox( _("Error: can't find file ") & cURL, 0, "TexMaths")

			ImportGraphicIntoClipboard = FALSE
			Exit Function

		End If

	End If

	' Import the graphic from URL into a new draw document
	Dim arg1(0) as New com.sun.star.beans.PropertyValue

	Dim oDrawDoc as Variant, oDrawDocCtrl as Variant
	 arg1(0).Name = "Hidden"
	 arg1(0).Value = TRUE

	oDrawDoc = StarDesktop.loadComponentFromURL( cURL_, "_blank", 0, arg1())
	oDrawDocCtrl = oDrawDoc.getCurrentController()
	
	' Get the draw page
	Dim oDrawPage as Variant, oImportShape as Variant
	oDrawPage = oDrawDoc.DrawPages(0)
	
	' Group shapes if SVG format because there are multiple shapes
	' PNG format has only one shape, so there is no need to group
	If sEqFormat = "svg" Then
	
		Dim oShapes as Variant
		oShapes = createUnoService("com.sun.star.drawing.ShapeCollection")

		Dim i as Integer
		For i = 0 To oDrawPage.getCount()-1
			oShapes.add(oDrawPage.getByIndex(i))
		Next
		
		oDrawPage.group(oShapes)

	End If

	' Get the shape
    oImportShape = oDrawPage(0) 

	Dim oImageSize as Variant
	Dim oShapeSize as Variant			
	oShapeSize = createUnoStruct("com.sun.star.awt.Size")

	' LibreOffice (or OpenOffice) version
	Dim sVersion as String
	sVersion = GetLOVersion()

	' If LibreOffice version >= 5.2 and < 6.1 and SVG format, 
	' reduce image size to mimic the text font size
    If Val(sVersion) >= 5.2 And Val(sVersion) < 6.1 And sEqFormat = "svg" Then
		
		' Get actual image size
		oImageSize = oImportShape.Size()
		
		' Set image size
		oShapeSize.Width = oImageSize.Width * 0.8
		oShapeSize.Height = oImageSize.Height * 0.8
		oImportShape.setSize(oShapeSize)

	End If

	' If PNG format, scale the image obtained from the dvipng external program
	If sEqFormat = "png" Then
		
		' Get actual image size, in pixels
		oImageSize = oImportShape.Graphic.SizePixel()
		
		' Set image size
		oShapeSize.Width = (oImageSize.Width * 35) * (72 / Val(sEqDPI))
		oShapeSize.Height = (oImageSize.Height * 35) * (72 / Val(sEqDPI))
		oImportShape.setSize(oShapeSize)

	End If	
	
	' Copy the image to clipboard and close the draw document
	oDrawDocCtrl.select(oImportShape)

	Dim oDispatcher as Variant
	Dim Array()
	oDispatcher = createUnoService( "com.sun.star.frame.DispatchHelper" )

	' Workaround for problem when exporting SVG to MS-Office
	If sEqFormat = "svg" Then
		oDispatcher.executeDispatch( oDrawDocCtrl.Frame, ".uno:ChangeBezier", "", 0, Array() )
	End If

	oDispatcher.executeDispatch( oDrawDocCtrl.Frame, ".uno:Copy", "", 0, Array() )
	oDrawDoc.close(True)
	oDrawDoc.dispose()
	
End Function


' Read the Latex attributes (parameters and text of the Latex equation) of the object
' Can read old TexMaths or ooolatex attributes
Function ReadAttributes( oShape as Variant ) as String

	Dim str as String

	On Error Resume Next

    ' Check if the object is an old TexMaths equation
	Dim oAttributes as Variant
    oAttributes = oShape.UserDefinedAttributes() 	
   	str = oAttributes.getByName("TexMathsArgs").Value  	
   	If str <> "" Then
   	
		ReadAttributes = str
		Exit Function

	Else
	
		' Check if the object is an ooolatex equation
		str = oAttributes.getByName("OOoLatexArgs").Value
   		
   		If str <> ""  Then
   		
			ReadAttributes = str
			Exit Function
   		
		Else   		
   		
   			' Read the image title
	   		str = oShape.Title
		
			' Attributes are stored in the image description
			If str = "TexMaths" Then
					
				ReadAttributes = oShape.Description
				Exit Function
				
			End If
							
		End If
		
		ReadAttributes = ""
				
	End If

End Function


' Write the Latex attributes (parameters and text of the Latex equation)
' into the object title and description
Sub SetAttributes( oShape as Variant, iEqSize as Integer, sEqType as String, sEqCode as String, sEqFormat as String, sEqDPI as String, sEqTransp as String, sEqName as String)

	oShape.Title = "TexMaths"
	oShape.Description = iEqSize & "�" & sEqType & "�" & sEqCode & "�" & sEqFormat & "�" & sEqDPI & "�" & sEqTransp & "�" & sEqName

End Sub



' Get the image size from the .dat file
Function GetImageSize() as com.sun.star.awt.Size

	Dim sFilePath as String
	Dim iNumber as Integer
	Dim sLine1 as String, sLine2 as String

	' Initializations
	iNumber = Freefile
	sLine1 = ""
	sLine2 = ""
	sFilePath = glb_TmpPath & "tmpfile.dat"

	' Read the .dat file
	If FileExists(sFilePath) Then
		
		Open sFilePath For Input As iNumber
			Line Input #iNumber, sLine1
			Line Input #iNumber, sLine2
		Close #iNumber

		Dim str1() as String, str2() as String
		Dim height as Double, width as Double, depth as Double

		' Image format is SVG
		If glb_Format = "svg" Then
			
			' Get the image width and height in mm
			str1 = Split(sLine2,"(")
			str2 = Split(str1(1),"mm")
			width = Val(str2(0))
			str1 = Split(sLine2,"x")
			str2 = Split(str1(2),"mm")
			height = Val(str2(0))

			' Convert image width and height to twips
			width = width*100
			height = height*100

		' Image format is PNG
		Else
	
			' Get the image depth, height and in mm
			str1 = Split(sLine2,"=")
			str2 = Split(str1(1)," ")
			depth = Val(str2(0))
			
			str1 = Split(sLine2,"=")
			str2 = Split(str1(2)," ")
			height = Val(str2(0))
	
			str1 = Split(sLine2,"=")
			str2 = Split(str1(3)," ")
			width = Val(str2(0))

			' Compute width and height (total height) in twips
			height = depth + height
			width = width*2.54/Val(glb_GraphicDPI)*1000
			height = height*2.54/Val(glb_GraphicDPI)*1000

		End If
	
		' Return image size		
		Dim oSize as Variant			
		oSize = createUnoStruct("com.sun.star.awt.Size")
		oSize.Width = width
		oSize.Height = height	
		GetImageSize = oSize

	' File not found
	Else

		MsgBox( _("Error: can't find file ") & sFilePath, 0, "TexMaths")
		Exit Function

	End If
	
End Function



' Get the vertical shift of the image according to the baseline position
' Return 0 if an error has occurred
Function GetVertShift() as Double

	Dim sFilePath as String
	Dim str1() as String, str2() as String
	Dim iNumber as Integer
	Dim sLine1 as String, sLine2 as String
		
	' Read the file that contains depth and height
	iNumber = Freefile
	sLine1 = ""
	sLine2 = ""
	
	sFilePath = glb_TmpPath & "tmpfile.bsl"
	If FileExists(sFilePath) Then
	
		Open sFilePath For Input As iNumber
			Line Input #iNumber, sLine1
			Line Input #iNumber, sLine2
		Close #iNumber
		
		' Get the depth and height
		Dim depth as Double, height as Double
		str1 = Split(sLine1,"=")
		str2 = Split(str1(1),"pt")
		depth = Val(str2(0))
		str1 = Split(sLine2,"=")
		str2 = Split(str1(1),"pt")
		height = Val(str2(0))
		
		' An error has occurred
		If depth + height = 0 Then
			
			ErrorDialog( _("LaTeX code was successfully compiled but equation image is empty, please check the equation syntax...") )
			GetVertShift = 0
		
		Else
	
			' Compute vertical shift and return value
			GetVertShift = height / (depth + height)
		
		End If
		
	' File not found
	Else

		MsgBox( _("Error: can't find file ") & sFilePath, 0, "TexMaths")
		Exit Function
	
	End If
	
End Function


' Display error on screen from file in temp directory 
Sub PrintError(sFile as String)

	Dim iNumber as Integer
	Dim sMsg as String, sLine as String
	If Not FileExists(glb_TmpPath & sFile) Then

		MsgBox( _("Error: the file ") & glb_TmpPath & sFile & _(" doesn't exist..."), 0, "TexMaths")
		Exit Sub

	End If

	iNumber = Freefile
	Open glb_TmpPath & sFile For Input As iNumber
 	While Not EOF(iNumber)
  		Line Input #iNumber, sLine
		sMsg = sMsg & sLine & chr(10)
	Wend
	Close #iNumber
	
	If sMsg = "" Then
		sMsg = _("An error has occurred, please check your TexMaths configuration...")
	End If
	
	ErrorDialog(sMsg)
	
End Sub



' Display file on screen from temp directory 
Sub PrintFile(sFile as String, sTitle as String)

	Dim iNumber as Integer
	Dim sMsg as String, sLine as String
	If Not FileExists(glb_TmpPath & sFile) Then

		MsgBox( _("Error: the file ") & glb_TmpPath & sFile & _(" doesn't exist..."), 0, "TexMaths")
		Exit Sub

	End If

	iNumber = Freefile
	Open glb_TmpPath & sFile For Input As iNumber
 	While Not EOF(iNumber)
  		Line Input #iNumber, sLine
		sMsg = sMsg & sLine & chr(10)
	Wend
	Close #iNumber
	
	If sMsg = "" Then
		sMsg = _("An error has occurred, please check your TexMaths configuration...")
	End If
	
	MessageDialog(sMsg, sTitle)
	
End Sub



' Convert decimal into two digits hexadecimal number as string
Function Hex2( Value as Integer) as String

	Dim Hex1 as String
	If Value = 0 Then 

		Hex2 = "00" 
		Exit Function

	End If

	Hex1 = Hex( Value )
	If Len( Hex1 ) = 1 Then Hex1 = "0" & Hex1
	
	Hex2 = Hex1

End Function


' Display given message in the status bar
Sub DisplayStatus(msg as String)

	glb_Status = msg

End Sub


' Add a slash if necessary
Function CheckPath( sPath as String) as String

	If Right(sPath,1) = "/" Then 
		CheckPath = sPath
	Else
		CheckPath = sPath & "/"
	End If

End Function


' Check if file exists and if not, displays an error message
Function CheckFile( sUrl as String, ErrorMsg as String) as Boolean

	If FileExists(sUrl) Then
		CheckFile = FALSE

	Else

		If ErrorMsg = "TexMaths" Then 
			ErrorMsg = _("Can't find ") & sUrl & chr(10) & _("Please check your installation...")
		End If
		MsgBox(ErrorMsg, 0, "TexMaths")
		CheckFile = TRUE

	End If

End Function


' Return TRUE if string "s" doesn't contains character "c" 
Function StringNotContains( s as String, c as String ) as Boolean

	StringNotContains = TRUE
	If (Len(s) <> 0) Then
		
		Dim j as Integer
		For j = 1 to Len(s)
			If Mid(s,j,1) = c Then 
				StringNotContains = FALSE
				Exit For
			End If
		Next

	End If

End Function 


' Return True if cPrefixString matches the beginning of cString (case sensitive)
' The following would return true...
'   IsPrefixString( "Jo", "John" )
'   IsPrefixString( "Jo", "Joseph" )
'   IsPrefixString( "Jo", "Jolly" )
' Copyright (c) 2003-2004 Danny Brewer 
Function IsPrefixString( ByVal cPrefixString As String, ByVal cString As String ) As Boolean
   IsPrefixString = (Left( cString, Len( cPrefixString ) ) = cPrefixString )
End Function


' Get access to the repository
Function GetConfigAccess(ByVal cNodePath as String,_
						 ByVal bWriteAccess as Boolean,_
						 Optional bEnableSync,_
						 Optional bLazyWrite ) as Variant
					 
	If IsMissing( bEnableSync ) Then bEnableSync = TRUE
	If IsMissing( bLazyWrite )  Then bLazyWrite = FALSE

	Dim oConfigProvider as Variant
	oConfigProvider = GetProcessServiceManager().createInstanceWithArguments(_
						"com.sun.star.configuration.ConfigurationProvider",_
						Array( MakePropertyValue( "enableasync", bEnableSync ) ) )

	Dim cServiceName as String
	If bWriteAccess Then
		cServiceName = "com.sun.star.configuration.ConfigurationUpdateAccess"
	Else
		cServiceName = "com.sun.star.configuration.ConfigurationAccess"
	EndIf

	Dim oConfigAccess as Variant
	oConfigAccess = oConfigProvider.createInstanceWithArguments(_
			cServiceName,_
      		Array(  MakePropertyValue( "nodepath",  cNodePath  ),_
					MakePropertyValue( "lazywrite", bLazyWrite ) ) )

	GetConfigAccess = oConfigAccess

End Function


' Create a PropertyValue structure from name and value pair
Function MakePropertyValue( Optional cName as String, Optional uValue as Variant) As com.sun.star.beans.PropertyValue

	' Create structure 
	Dim oPropertyValue as Variant
	oPropertyValue = createUnoStruct( "com.sun.star.beans.PropertyValue" )
	
	' Set name and value pair
	If Not IsMissing( cName )  Then oPropertyValue.Name  = cName
	If Not IsMissing( uValue ) Then oPropertyValue.Value = uValue

	' Return structure
	MakePropertyValue = oPropertyValue

End Function


' On Windows, generate path as "C:\path_to_file\"
Function WinPath(sPath as String) as String

	sPath = ConvertFromUrl(sPath)
	WinPath = """" & sPath &  """"

End function


' Determine document type from the services that are supported
' Author Andrew Pitonyak
Function GetDocumentType( oDoc as Variant) as String

	Dim sImpress as String, sCalc as String, sDraw as String, sBase as String, sMath as String, sWriter as String

	sCalc    = "com.sun.star.sheet.SpreadsheetDocument"
	sImpress = "com.sun.star.presentation.PresentationDocument"
	sDraw    = "com.sun.star.drawing.DrawingDocument"
	sBase    = "com.sun.star.sdb.DatabaseDocument"
	sMath    = "com.sun.star.formula.FormulaProperties"
	sWriter  = "com.sun.star.text.TextDocument"

	On Local Error GoTo NO_DOCUMENT_TYPE
	
	If oDoc.SupportsService(sCalc) Then
  		GetDocumentType = "scalc"
	ElseIf oDoc.SupportsService(sWriter) Then
		GetDocumentType = "swriter"
	ElseIf oDoc.SupportsService(sDraw) Then
		GetDocumentType = "sdraw"
	ElseIf oDoc.SupportsService(sMath) Then
		GetDocumentType = "smath"
	ElseIf oDoc.SupportsService(sImpress) Then
		GetDocumentType = "simpress"
	ElseIf oDoc.SupportsService(sBase) Then
		GetDocumentType = "sbase"
	End If

	NO_DOCUMENT_TYPE:

	If Err <> 0 Then

  		GetDocumentType = ""
  		Resume GO_ON
  		GO_ON:

	End If

End Function



' Get application locale
' Original author : Laurent Godard
' e-mail : listes.godard@laposte.net
' Modified to return the complete locale string
' (not only the first two characters) 
Function GetLocale() as string

	Dim oSet as Variant, oConfigProvider as Variant
	Dim oParm(0) As New com.sun.star.beans.PropertyValue
	Dim sProvider as String, sAccess as String
	
	sProvider = "com.sun.star.configuration.ConfigurationProvider"
	sAccess   = "com.sun.star.configuration.ConfigurationAccess"
	
	oConfigProvider = createUnoService(sProvider)
	oParm(0).Name = "nodepath"
	oParm(0).Value = "/org.openoffice.Setup/L10N"
	oSet = oConfigProvider.createInstanceWithArguments(sAccess, oParm())

	GetLocale = oSet.getbyname("ooLocale")

End Function



' Translation function
' Replace each string like _("string example") with its translation
' Original author : Pierre Chef, june 2009
' Available under the terms of the WTFPL
Function _(msgid as String) as String

	Dim i as Integer
	Dim sTrans as String

	' Read the appropriate po file at the first time
	If glb_PoFileRead <> 1 then ReadPoFile(glb_PkgPath & "po")

	' Look for the corresponding translated string
	For i = 0 to Ubound(glb_MsgId)
	
		If glb_MsgId(i) = msgid Then
       		
       		sTrans = glb_MsgStr(i)
      		Exit For
      			
		End If

	Next i

	' Return the translated string
	If sTrans = "" Then sTrans = msgid
	_() = sTrans

End Function



' Read po file according to the current locale
' and construct the tables used to store translated strings
' Original author: Pierre Chef, june 2009
' Available under the terms of the WTFPL
Sub ReadPoFile(podir as string)

	Dim oFileAccess as Variant
	Dim sline as String        ' Line read
	Dim lineLen as Integer     ' Length of line read
	Dim msgCounter as Integer  ' Message counter
	Dim position as Integer    ' Position in a string
	Dim quotePos as Integer    ' Position after first quotation mark
	Dim transType as String    ' Translation type : msgid or msgstr
	Dim message as String      ' String contained in msgid or msgstr
	Dim locale as String       ' Locale code : fr, en, es, it, de, ...
	Dim pofile as String       ' po file path
   
	' Simple file access object
	oFileAccess = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
   
	' The podir must be a folder
	If oFileAccess.isFolder(podir) Then

		' Add eventually a trailing slash
		If right(podir,1) <> "/" then podir = podir + "/"      
      
      	' Get current locale
      	locale = GetLocale()
      
		' Po file path      
		pofile = podir + locale + ".po"

		' File must exist
		If oFileAccess.exists(pofile) Then
	 
			' Open po file in the same way as ReadTextFileUtf8()
			' because we may have UTF-8 characters within it
			Dim oTextFile as Variant, oFileStream as Variant
			oFileStream = oFileAccess.openFileRead(pofile)
			oTextFile = createUnoService("com.sun.star.io.TextInputStream")
			oTextFile.InputStream = oFileStream
		
		Else
         	Exit Sub
		End If
	
	Else
		Exit Sub
	End If
   
	' Initialize counter
	msgCounter = -1

	' Read the po file, line by line
	While Not oTextFile.IsEOF
 
  		' Read line
  		sline = oTextFile.readLine
      	lineLen=len(sline)
      
     	quotePos = InStr(sline,"""")+1

		If lineLen > 2 Then
			
			Select Case left(sline,1)
				
				' Po file comment
				Case "#":
				
				' Identifier
				Case "m":
            		
            		If left(sline,5) = "msgid" Then
               			                 			
                		transType = "id"
                		msgCounter = msgCounter+1
                		redim Preserve glb_MsgId(msgCounter)
                		redim Preserve glb_MsgStr(msgCounter)
               								
						message = mid(sline,quotePos,lineLen-quotePos)
 					
 					Elseif left(sline,6) = "msgstr" Then
 						
 						transType="str"
               			               
               			message = mid(sline,quotePos,lineLen-quotePos)
            
            		Endif
            
            		' Update tables
            		UpdTransTables(message,msgCounter,transType)
				
				' String (quotePos=1 obviously)
				Case """":
            		
            		' Update tables
            		message = mid(sline,quotePos,lineLen-quotePos)
            		UpdTransTables(message,msgCounter,transType)
         		
         		' Other :  error in the po file
         		Case Else
            						
				End Select
      
		End If
	
	Wend ' End file reading loop
	
	' Close file
	oFileStream.closeInput
	oTextFile.closeInput
	
	' Set flag
	glb_PoFileRead = 1

End Sub


' Update the tables used to store translated strings
' Original author: Pierre Chef, june 2009
' Available under the terms of the WTFPL
Sub UpdTransTables(message as String, msgCounter as Integer, transType as String)
   
	If message = "" Then Exit Sub
	
	If transType = "id" Then
		glb_MsgId(msgCounter)= glb_MsgId(msgCounter) + message

	Elseif transType = "str" Then
		glb_MsgStr(msgcounter)= glb_MsgStr(msgcounter) + message
	
	Endif

End Sub


' Put the content of the clipboard into a string
' Original author: Andrew Pitonyak
Function ClipboardToText() as String

	On Error Goto ErrorHandler ' Enables error handling

	Dim oClip as Variant, oClipContents as Variant, oTypes as Variant, oConverter as Variant
	Dim i as Integer, iPlainLoc as Integer
	Dim sContent as String


	oClip = createUnoService( "com.sun.star.datatransfer.clipboard.SystemClipboard" )
	oConverter = createUnoService( "com.sun.star.script.Converter" )
	oClipContents = oClip.getContents()
	oTypes = oClipContents.getTransferDataFlavors()

	iPlainLoc = -1
	For i = LBound(oTypes) To UBound(oTypes)

		If oTypes(i).MimeType = "text/plain;charset=utf-16" Then

			iPlainLoc = i
			Exit For

		End If

	Next

	sContent = ""
	If iPlainLoc >= 0 Then

		Dim oData as Variant
		oData = oClipContents.getTransferData(oTypes(iPlainLoc))
		sContent = oConverter.convertToSimpleType(oData, com.sun.star.uno.TypeClass.STRING)

	End If

	ClipboardToText = sContent
	Exit Function

' Handle error that sometimes occurs in oClipContents.getTransferData()
ErrorHandler:

	'MsgBox( "Error in ClipboardToText()", 0, "TexMaths")
	'MsgBox( "Error code: " + Err + Chr$(13) + Error$, 0, "TexMaths")
	
	ClipboardToText = sContent

End Function



' Put a string content into the clipboard
' Original author : DannyB
Sub TextToClipboard( sContent As String )
   
	Dim oDoc as Variant, oText as Variant, oCursor as Variant
	
	' Create an empty hidden Writer document
	oDoc = StarDesktop.loadComponentFromURL( "private:factory/swriter", "_blank", 0, Array( MakePropertyValue( "Hidden", True ) ) )
	
	' Get the text of the document
	oText = oDoc.getText()
	
	' Get a cursor that can move over or to any part of the text
	oCursor = oText.createTextCursor()
	
	' Insert text and paragraph breaks into the text, at the cursor position
	oText.insertString( oCursor, sContent, False )
	
	' Dispatch commands
	Dim oFrame as Variant, oDispatcher as Variant
	oFrame = oDoc.CurrentController.Frame
	oDispatcher = createUnoService( "com.sun.star.frame.DispatchHelper" )
	oDispatcher.executeDispatch(oFrame,".uno:SelectAll","",0,Array())
	oDispatcher.executeDispatch(oFrame,".uno:Copy","",0,Array())

	' Close document
	oDoc.close( True )

End Sub 


' Position the cursor at the most left position of the selection
' Author: Andrew Pitonyak
' email:   andrew@pitonyak.org
' oSel is a text selection or cursor range
Function GetLeftMostCursor(oSel as Variant) as Variant

	Dim oRange as Variant    ' Right most range
	Dim oCursor as Variant   'Cursor at the right most range
	
	If oSel.getText().compareRegionStarts(oSel.getEnd(), oSel) >= 0 Then
		oRange = oSel.getEnd()
	Else
		oRange = oSel.getStart()
	End If
	
	oCursor = oSel.getText().CreateTextCursorByRange(oRange)
	oCursor.goRight(0, False)
	GetLeftMostCursor = oCursor

End Function


' Position the cursor at the most right position of the selection
' Author: Andrew Pitonyak
' email:   andrew@pitonyak.org
' oSel is a text selection or cursor range
Function GetRightMostCursor(oSel as Variant) as Variant

	Dim oRange as Variant    ' Right most range
	Dim oCursor as Variant   'Cursor at the right most range

	If oSel.getText().compareRegionStarts(oSel.getEnd(), oSel) >= 0 Then
	  oRange = oSel.getStart()
	Else
	  oRange = oSel.getEnd()
	End If

	oCursor = oSel.getText().CreateTextCursorByRange(oRange)
	oCursor.goLeft(0, False)
	GetRightMostCursor = oCursor

End Function



' Test if a component (Writer, Draw, Impress) is installed
Function ComponentInstalled( sName as String ) as Boolean

	Dim oModuleManager as Variant
	oModuleManager = CreateUnoService( "com.sun.star.frame.ModuleManager" )
    
	ComponentInstalled = FALSE

    If (sName = "Writer") and oModuleManager.hasByName( "com.sun.star.text.TextDocument" ) Then
		ComponentInstalled = TRUE
	End If
		
	If (sName = "Impress") and  oModuleManager.hasByName( "com.sun.star.presentation.PresentationDocument" ) Then
		ComponentInstalled = TRUE
	End If

	If (sName = "Draw") and oModuleManager.hasByName( "com.sun.star.drawing.DrawingDocument" ) Then
		ComponentInstalled = TRUE
	End If
 
End Function


' Transfer animation from an old shape to a new shape in Impress mode
' Author: Daniel Fett
Sub TransferAnimations(slide as Variant, original as Variant, replacement as Variant)

    Dim oMainSequence as Variant, oClickNodes as Variant, oClickNode as Variant, oGroupNodes as Variant
    Dim oGroupNode as Variant, oEffectNodes as Variant, oEffectNode as Variant, oAnimNodes as Variant, oAnimNode as Variant

    oMainSequence = GetMainSequence(slide)    
	If oMainSequence = null Then ' Exit if null object
		Exit Sub
  	End if

    oClickNodes = oMainSequence.createEnumeration()

    While oClickNodes.hasMoreElements()

        oClickNode = oClickNodes.nextElement()
        oGroupNodes = oClickNode.createEnumeration()

        While oGroupNodes.hasMoreElements()

            oGroupNode = oGroupNodes.nextElement()
            oEffectNodes = oGroupNode.createEnumeration()

            While oEffectNodes.hasMoreElements()

                oEffectNode = oEffectNodes.nextElement()
                oAnimNodes = oEffectNode.createEnumeration()

                While oAnimNodes.hasMoreElements()

                    oAnimNode = oAnimNodes.nextElement()
                    If EqualUnoObjects(original, oAnimNode.target) Then
                    	oAnimNode.target = replacement
                    End If 

                Wend
            
            Wend
    
        Wend
    
    Wend

End Sub

 
 
' Get the main sequence from the given page
' Author: Daniel Fett
Function GetMainSequence(oPage as Variant) as Variant

	Dim oMainSeq as Integer, oNodes as Variant, oNode as Variant
    oMainSeq = com.sun.star.presentation.EffectNodeType.MAIN_SEQUENCE
 
 	' Initialize to null
	GetMainSequence = null

    oNodes = oPage.AnimationNode.createEnumeration()
    
    While oNodes.hasMoreElements()
        
        oNode = oNodes.nextElement()
        
        If GetNodeType(oNode) = oMainSeq Then
            GetMainSequence = oNode
            Exit Function
        End If
    
    Wend

End Function


' Get the type of a node
' Author: Daniel Fett
Function GetNodeType(oNode as Variant) as Integer

	Dim oData as Variant
	
    For each oData in oNode.UserData

        If oData.Name = "node-type" Then
            GetNodeType = oData.Value
            Exit Function
        End If

    Next oData

End Function


' Return all strings that are contained in a given LaTeX command. 
' For example, for the source code "\test{a} b c" and the command (search string) "test", it will return an array with one value, namely "a".
' Author: Daniel Fett
Function FindInLatexCommand(sSourceCode as String, sCommand as String)

	Dim aStrings() As String
	Dim sSearch As String
	Dim iLast As Integer
	Dim iCommandStart As Integer

	iCommandStart = 1
	iLast = 1
	sSearch = "\" & sCommand & "{"
	
	Do While iLast <> 0 and iCommandStart <> 0

		iCommandStart = InStr(iLast, sSourceCode, sSearch)

		If iCommandStart <> 0 Then

			' We found the beginning of the command. Great, now search for the end!
			iLast = InStr(iCommandStart, sSourceCode, "}")
			
			Dim n as Integer
			n = UBound(aStrings) + 1 ' To prevent a crash in Openoffice 4.x
			ReDim Preserve aStrings(n)
			aStrings(UBound(aStrings)) = Mid(sSourceCode, iCommandStart + Len(sSearch), iLast - (iCommandStart + Len(sSearch)))

		End If

	Loop
	FindInLatexCommand = aStrings

End Function
	


' Returns TRUE if anything is selected in a Writer document
' Based on a function written by Andrew Pitonyak
Function IsAnythingSelected(oDoc as Variant) As Boolean

	Dim oSels as Variant   ' All of the selections
	Dim oSel as Variant    ' A single selection
	Dim oCursor as Variant 'A temporary cursor

	IsAnythingSelected = FALSE

	If IsNull(oDoc) Then Exit Function

	' The current selection in the current controller
	' If there is no current controller, it returns NULL
	oSels = oDoc.getCurrentSelection()
	If IsNull(oSels) Then Exit Function

	' Text
	If oSels.getImplementationName() = "SwXTextRanges" Then
	
		' I have never seen a selection count of zero
		If oSels.getCount() = 0 Then Exit Function
		
		' If there are multiple selections, then assume something is selected
		If oSels.getCount() > 1 Then
			
			IsAnythingSelected = TRUE
		
		Else
		
			' If only one thing is selected, however, then check to see
			' if the selection is collapsed. In other words, see if the
			' end location is the same as the starting location.
			' Notice that I use the text object from the selection object
			' because it is safer than assuming that it is the same as the
			' documents text object.
			oSel = oSels.getByIndex(0)
			oCursor = oSel.getText().CreateTextCursorByRange(oSel)
			If Not oCursor.IsCollapsed() Then IsAnythingSelected = TRUE
		
		End If
	
	' Tables, graphic elements, images, etc.
	Else
		
		IsAnythingSelected = TRUE
	
	End If

End Function
