Function MakeEquation(iEqSize as Integer, sEqType as String, sEqCode as String, sEqFormat as String, sEqDPI as String, sEqTransp as String, sEqName as String, oShape as Variant) as Integer

	Dim iNumber as Integer
	Dim cURL as String, sShellArg as String, sMsg as String, sLatexCode as String, sShellCommand as String
	Dim sVersion as String

	' LibreOffice (or OpenOffice) major version
	sVersion = Left(GetLOVersion(),1)

	' Save initial clipboard content because it will be lost when pasting the image
	If Not (getGUIType() = 1 And sVersion = "5") Then ' Doesn't work on Windows with Libreoffice 5.x
	
		Dim sClipContent as String
		sClipContent = ClipboardToText()
	
	End If

	' We first test if there is LateX code
	If sEqCode = "" Then 

		MsgBox( _("Please enter an equation..."), 0, "TexMaths")
		MakeEquation = -1
		Exit Function

	End If

	' Some environments only work in LaTeX mode 
	If IsPrefixString("\begin{align",sEqCode) Or IsPrefixString("\begin{eqnarray",sEqCode) Or IsPrefixString("\begin{gather",sEqCode) _
	Or isPrefixString("\begin{flalign",sEqCode) Or IsPrefixString("\begin{multline",sEqCode) Or IsPrefixString("\begin{minted",sEqCode) Then
		sEqType = "latex" 
	End If

	' Check if the LaTeX code has any dependencies that can be fulfilled easily by copying files from source dir to tmp dir
	' We check for include and input (with .tex files) and usepackage (with .sty files)
	' Note that we search in the preamble and in the latex code (we concatenate the two)
	' Patch by Daniel Fett

	' Document URL
	Dim sUrl as String
	sUrl = ThisComponent.getURL()
	
	' Open service file and an output stream
	Dim oFileAccess as Variant, oTextStream as Variant
	oFileAccess = createUnoService("com.sun.star.ucb.SimpleFileAccess")
	oTextStream  = createUnoService("com.sun.star.io.TextOutputStream")

	Dim sPotentialFile as String
	Dim sFilePath as String 
	
	Dim aSplittedURL() as String
	Dim sCurrentDir as String
	
	aSplittedURL = split(sUrl, "/")

	For Each sPotentialFile in FindInLatexCommand(glb_Preamble & sEqCode, "include")

		' Document not saved
		If (sUrl = "") Then	
			MsgBox( _("Please save the document before using an \include command..."), 0, "TexMaths")
			MakeEquation = -1
			Exit Function
		End If
		
		' When using an \include command, the equation type must be set to "latex"
		sEqType = "latex" 	
		oDlgMain.getControl("TypeLatex").setState(1)

		' Copy tex file to the tmp directory
		sCurrentDir = Left(sUrl, Len(sUrl) - Len(aSplittedUrl(UBound(aSplittedUrl))))
		sFilePath = sCurrentDir & "/" & sPotentialFile & ".tex"				
		If oFileAccess.exists(sFilePath) Then
			oFileAccess.copy(sFilePath, ConvertToURL( glb_TmpPath & sPotentialFile & ".tex"))
		End If

	Next
	
	For Each sPotentialFile in FindInLatexCommand(glb_Preamble & sEqCode, "input")

		' Document not saved
		If (sUrl = "") Then
			MsgBox( _("Please save the document before using an \input command..."), 0, "TexMaths")
			MakeEquation = -1
			Exit Function
		End If

		' Copy tex file to the tmp directory
		sCurrentDir = Left(sUrl, Len(sUrl) - Len(aSplittedUrl(UBound(aSplittedUrl))))
		sFilePath = sCurrentDir & "/" & sPotentialFile & ".tex"
		If oFileAccess.exists(sFilePath) Then
			oFileAccess.copy(sFilePath, ConvertToURL( glb_TmpPath & sPotentialFile & ".tex"))
		End If
	Next
	
	For Each sPotentialFile in FindInLatexCommand(glb_Preamble & sEqCode, "usepackage")
	
		' Document not saved
		If (sUrl = "") Then
	
			sFilePath = ConvertToURL( glb_TmpPath & sPotentialFile & ".sty") ' Path of sty file in tmp dir
			If oFileAccess.exists(sFilePath) Then
				oFileAccess.kill(sFilePath)
			End If
		
		Else

			sCurrentDir = Left(sUrl, Len(sUrl) - Len(aSplittedUrl(UBound(aSplittedUrl))))
			sFilePath = sCurrentDir & "/" & sPotentialFile & ".sty"
			If oFileAccess.exists(sFilePath) Then
				oFileAccess.copy(sFilePath, ConvertToURL( glb_TmpPath & sPotentialFile & ".sty"))
			End If
		
		End If

	Next


	' Build the LaTeX code, depending on the selected mode
	If sEqType = "inline"  Then
		
		sLatexCode = "\begin{document}" & chr(10) &_
		"\newsavebox{\eqbox}" & chr(10) &_
		"\newlength{\width}" & chr(10) &_
		"\newlength{\height}" & chr(10) &_
		"\newlength{\depth}" & chr(10) & chr(10) &_
		"\begin{lrbox}{\eqbox}" & chr(10) &_
		"{$ " & sEqCode & " $}" & chr(10) &_
		"\end{lrbox}" & chr(10) & chr(10) &_
		"\settowidth {\width}  {\usebox{\eqbox}}" & chr(10) &_
		"\settoheight{\height} {\usebox{\eqbox}}" & chr(10) &_
		"\settodepth {\depth}  {\usebox{\eqbox}}" & chr(10) &_
		"\newwrite\file" & chr(10) &_
		"\immediate\openout\file=\jobname.bsl" & chr(10) &_
		"\immediate\write\file{Depth = \the\depth}" & chr(10) &_
		"\immediate\write\file{Height = \the\height}" & chr(10) &_
		"\addtolength{\height} {\depth}" & chr(10) &_
		"\immediate\write\file{TotalHeight = \the\height}" & chr(10) &_
		"\immediate\write\file{Width = \the\width}" & chr(10) &_
		"\closeout\file" & chr(10) &_
		"\usebox{\eqbox}" & chr(10) &_
		"\end{document}" & chr(10)
		
	ElseIf sEqType = "display"  Then
		
		sLatexCode = "\begin{document}" & chr(10) &_
		"\newsavebox{\eqbox}" & chr(10) &_
		"\newlength{\width}" & chr(10) &_
		"\newlength{\height}" & chr(10) &_
		"\newlength{\depth}" & chr(10) & chr(10) &_
		"\begin{lrbox}{\eqbox}" & chr(10) &_
		"{$\displaystyle " & sEqCode & " $}" & chr(10) &_
		"\end{lrbox}" & chr(10) & chr(10) &_
		"\settowidth {\width}  {\usebox{\eqbox}}" & chr(10) &_
		"\settoheight{\height} {\usebox{\eqbox}}" & chr(10) &_
		"\settodepth {\depth}  {\usebox{\eqbox}}" & chr(10) &_
		"\newwrite\file" & chr(10) &_
		"\immediate\openout\file=\jobname.bsl" & chr(10) &_
		"\immediate\write\file{Depth = \the\depth}" & chr(10) &_
		"\immediate\write\file{Height = \the\height}" & chr(10) &_
		"\addtolength{\height} {\depth}" & chr(10) &_
		"\immediate\write\file{TotalHeight = \the\height}" & chr(10) &_
		"\immediate\write\file{Width = \the\width}" & chr(10) &_
		"\closeout\file" & chr(10) &_
		"\usebox{\eqbox}" & chr(10) &_
		"\end{document}" & chr(10)

	ElseIf sEqType = "latex" Then
	
		sLatexCode = sEqCode
		
	End If

	' Create the LaTeX file with the LatexCode
	cURL = ConvertToURL( glb_TmpPath & "tmpfile.tex" )
	If oFileAccess.exists( cURL ) Then oFileAccess.kill( cURL )
    oTextStream.setOutputStream(oFileAccess.openFileWrite(cURL))
	
	If sEqType = "latex" Then
	
		If glb_IgnorePreamble = TRUE Then
	
			oTextStream.writeString(sLatexCode)
	
		Else
	
			oTextStream.writeString( _
			    	"\documentclass[10pt,dvips]{article}" & chr(10) &_
		    		glb_Preamble & chr(10) & chr(10) &_
			    	"\pagestyle{empty}" & chr(10) &_
			    	"\begin{document}" & chr(10) &_
			    	sLatexCode & chr(10) &_
			    	"\end{document}" )
	
		End If

	Else
	
		oTextStream.writeString( _
		    	"\documentclass[10pt,dvips]{article}" & chr(10) &_
		    	glb_Preamble & chr(10) & chr(10) &_
		    	"\pagestyle{empty}" & chr(10) &_
		    	sLatexCode )
	
	End If

	' Close the file
    oTextStream.closeOutput()
    
    ' Test the existence of the LaTeX file...
    If CheckFile( glb_TmpPath & "tmpfile.tex" , _
    		_("The file ") & ConvertFromURL(glb_TmpPath) & _("tmpfile.tex can't be created") & chr(10) & _
			_("Please check your installation...") ) Then 

		ConfigDialog()
		MakeEquation = -1
		Exit Function

	End If
	
	' Windows
	If getGUIType() = 1 Then

		sShellCommand = ConvertToURL( GetScriptPath() )
		sShellArg = sEqFormat & " "  & sEqDPI & " "  & sEqTransp & " " & glb_TmpPath & " " & glb_Compiler

	' Linux or Mac OS X
	Else 					
		
		sShellCommand = "/bin/sh"
		sShellArg = "'" & ConvertFromURL(GetScriptPath()) & "' " & sEqFormat &_
				    " " & sEqDPI & " " & sEqTransp & " '"  & ConvertFromURL(glb_TmpPath) & "' " & glb_Compiler
	End If

    ' Remove Latex output file
	cURL = ConvertToURL( glb_TmpPath & "tmpfile.out" )
	If oFileAccess.exists( cURL ) Then oFileAccess.kill( cURL )

	' Call the script
	Shell(sShellCommand, 2, sShellArg, TRUE)

	' Check the result
	Dim sDviFile as String
	If glb_Compiler = "latex" Then
		sDviFile = "tmpfile.dvi"
	Else
		sDviFile = "tmpfile.xdv"
	End If
	 
 	If Not FileExists(glb_TmpPath & sDviFile) _
 	   and Not FileExists(glb_TmpPath & "tmpfile.out")  Then
		MsgBox( _("No file created in the directory:") & _
  		       chr(10) & ConvertFromURL(glb_TmpPath), 0, "TexMaths")
  		MakeEquation = -1
  		Exit Function

 	ElseIf Not FileExists(glb_TmpPath & sDviFile) _
 	       and FileExists(glb_TmpPath & "tmpfile.out")  Then ' Latex error

  		PrintError("tmpfile.out")
  		MakeEquation = -1
  		Exit Function

 	ElseIf CheckFile(glb_TmpPath & "tmpfile." & sEqFormat,_
   		_("Script error: the dvi file was not converted to ") & sEqFormat & "! " & chr(10) & chr(10) &_
   		_("Please check your system configuration...") ) Then
   		
	  		MakeEquation = -1
	  		Exit Function

 	End If


	' Create the Controller and dispatcher for current document
	Dim ret as Boolean
	Dim oDoc as Variant, oDocCtrl as Variant, oDispatcher as Variant, oGraphic as Variant, oShapeSize as Variant
	oDoc = ThisComponent
	oDocCtrl = oDoc.getCurrentController()
	oDispatcher = createUnoService("com.sun.star.frame.DispatchHelper")


	' ================== Current document is a Writer document ==================
	If GetDocumentType(oDoc) = "swriter" Then

		Dim oViewCursor as Variant, oCursor as Variant
		Dim AnchorType as Integer

		' If there is already an equation image, remove it
		If EditEquation Then

			' Select image (ensuring compatibility with previous TexMaths versions)
			Dim oSelection as Variant
			On Error Goto SelectionError
			oSelection = oDocCtrl.getSelection().GetByIndex(0)		
			
			' Get selected image anchor
			Dim oAnchor as Variant
			oAnchor = oSelection.getAnchor()
			AnchorType = oSelection.AnchorType

			' Unselect image
		 	oDispatcher.executeDispatch( oDocCtrl.Frame, ".uno:Escape", "", 0, Array())
		 	
		 	' Workaround for an issue with gtk3 backend
		 	If isNull(oShape) Then
		   	  	oDoc.drawPage.remove(oSelection)
			Else
				oDoc.drawPage.remove(oShape)
			End If
		 
		 End If
		 
		' Set vertical alignement to middle when Word compatibility is requested
		If glb_WordVertAlign = TRUE Then
			
			dim args4(0) as new com.sun.star.beans.PropertyValue
			args4(0).Name = "VerticalParagraphAlignment"
			args4(0).Value = 3		
			oDispatcher.executeDispatch(oDocCtrl.Frame, ".uno:VerticalParagraphAlignment", "", 0, args4())
		
		End If
		
		' Set text cursor position to the view cursor position
		oViewCursor = oDocCtrl.ViewCursor
		oCursor = oViewCursor.Text.createTextCursorByRange(oViewCursor)
		oCursor.gotoRange(oViewCursor,FALSE)
								
		' Import the new image into the clipboard
   		ret = ImportGraphicIntoClipboard(ConvertToURL( glb_TmpPath & "tmpfile."& sEqFormat), sEqFormat, sEqDPI, sEqTransp)
   		If ret = FALSE Then
   		
	  		MakeEquation = -1
	  		Exit Function
   		
   		 End If
	
		' Paste image to the current document
		oDispatcher.executeDispatch( oDocCtrl.Frame, ".uno:Paste", "", 0, Array())
	
	  	' Select image
	   	oGraphic = oDocCtrl.getSelection().GetByIndex(0)

		' Set the graphic object name
		oGraphic.Name = sEqname

		' Scale image
		oShapeSize = oGraphic.Size
		oShapeSize.Width = oShapeSize.Width * (iEqSize / 10)
		oShapeSize.Height = oShapeSize.Height * (iEqSize / 10)
		oGraphic.Size = oShapeSize	

		' In edit mode, anchor the image as it was
		If EditEquation Then
		
			Select Case AnchorType

				Case ANCHOR_TO_PARA
					oDispatcher.executeDispatch( oDocCtrl.Frame, ".uno:SetAnchorToPara", "", 0, Array())
					oGraphic.setPosition(oImgPosition)
				
				Case ANCHOR_TO_CHAR
					oDispatcher.executeDispatch( oDocCtrl.Frame, ".uno:SetAnchorToChar", "", 0, Array())
					' Don't position image in this case
				
				Case ANCHOR_TO_PAGE
					oDispatcher.executeDispatch( oDocCtrl.Frame, ".uno:SetAnchorToPage", "", 0, Array())
					oGraphic.setPosition(oImgPosition)					
				
				Case ANCHOR_AT_CHAR
					oDispatcher.executeDispatch( oDocCtrl.Frame, ".uno:SetAnchorAtChar", "", 0, Array())
					oGraphic.setPosition(oImgPosition)
			
			End Select
		
		' New equations are anchored to char
		Else
		
			AnchorType = ANCHOR_TO_CHAR
			oDispatcher.executeDispatch( oDocCtrl.Frame, ".uno:SetAnchorToChar", "", 0, Array())
			
		End If
	
		 ' Set some image properties
		oGraphic.TopMargin = 0
		oGraphic.BottomMargin = 0
		oGraphic.VertOrient = 0

		' Adjust vertical position for Inline or Display equations when image is anchored to char
		If AnchorType = ANCHOR_TO_CHAR and ( sEqtype = "inline" or sEqType = "display" ) Then
						
			' Get vertical shift coefficient
			Dim coef as Double
			coef = getVertShift()
					
			' Adjust the vertical position to the baseline
			oGraphic.VertOrientPosition = - coef * oShapeSize.Height

		End If

		' Set image attributes (size, type LaTeX code) for further editing
		SetAttributes( oGraphic, iEqSize, sEqType, sEqCode, sEqFormat, sEqDPI, sEqTransp, sEqName )

		' Save the paragraph style
		Dim ParaStyleName as String
		ParaStyleName = oViewCursor.ParaStyleName

		' Trick to ensure the image is located at the cursor position
		' => cut image, position cursor, then paste
		' => otherwise, the image can be anywhere within the paragraph 
		oDispatcher.executeDispatch(oDocCtrl.Frame, ".uno:Cut", "", 0, Array())
		oViewCursor.gotoRange(oCursor, FALSE)
		oDispatcher.executeDispatch(oDocCtrl.Frame, ".uno:Paste", "", 0, Array())

		' Restore paragraph style if it changed
		If ParaStyleName <> oViewCursor.ParaStyleName Then

			oViewCursor.ParaStyleName = ParaStyleName

		End If

		' Deselect image
		oDispatcher.executeDispatch(oDocCtrl.Frame, ".uno:Escape", "", 0, Array())


	' ================== Current document is an Impress or Draw document ==================
	ElseIf GetDocumentType(oDoc) = "simpress" or GetDocumentType(oDoc) = "sdraw" Then
		
 		' Edit equation: remove old image
 		If EditEquation Then 
 			
 			' Fill in reference to original shape (used to keep animations and Z index)
			Dim oOriginalShape as Variant
			Dim oOriginalShapeZOrder as Long
			oOriginalShape = oDocCtrl.getSelection().getByIndex(0)
			oOriginalShapeZOrder = oOriginalShape.ZOrder
 			oDispatcher.executeDispatch(oDocCtrl.Frame,".uno:Cut","", 0, Array() )
 		
 		End If
		
		' Import the new image to the clipboard
   		ret = ImportGraphicIntoClipboard(ConvertToURL( glb_TmpPath & "tmpfile." & sEqFormat), sEqFormat, sEqDPI, sEqTransp)
   		If ret = FALSE Then
	  		MakeEquation = -1
  			Exit Function
   		End If

		' Paste image to the current document
		oDispatcher.executeDispatch( oDocCtrl.Frame, ".uno:Paste", "", 0, Array())

		' Select image
		oGraphic = oDocCtrl.getSelection().getByIndex(0)

		' Set the graphic object name
		oGraphic.Name = sEqname

		' Edit equation: set its position equal to the previous one
		If EditEquation Then
			oGraphic.setPosition(oShapePosition)
		
		' New equation: Position the image at the center of the visible area
		Else
		
			Dim oPosition as Variant
			oPosition = createUnoStruct( "com.sun.star.awt.Point" ) 
			oPosition.X = oDocCtrl.VisibleArea.X + oDocCtrl.VisibleArea.Width / 2 - (oGraphic.Size.Width*iEqSize/10) / 2
			oPosition.Y = oDocCtrl.VisibleArea.Y + oDocCtrl.VisibleArea.Height / 2	- (oGraphic.Size.Height* iEqSize/10) / 2
			oGraphic.setPosition(oPosition)

		End If

		' Scale the image
		oShapeSize = oGraphic.Size
		oShapeSize.Width = oShapeSize.Width * iEqSize / 10
		oShapeSize.Height = oShapeSize.Height * iEqSize / 10
		oGraphic.Size = oShapeSize
		
		' Set image attributes (size, type LaTeX code) for further editing
		SetAttributes( oGraphic, iEqSize, sEqType, sEqCode, sEqFormat, sEqDPI, sEqTransp, sEqName)

		' Trick to allow undoing the equation insertion
		oDispatcher.executeDispatch( oDocCtrl.Frame, ".uno:Cut", "", 0, Array())
		oDispatcher.executeDispatch( oDocCtrl.Frame, ".uno:Paste", "", 0, Array())

		' Select image
		oGraphic = oDocCtrl.getSelection().getByIndex(0)

		' Edit equation: if in Normal view mode, transfer animations from old shape to new shape
 		If EditEquation and GetDocumentType(oDoc) = "simpress" and oDocCtrl.DrawViewMode = 0 Then
			TransferAnimations(oDocCtrl.getCurrentPage(), oOriginalShape, oGraphic) 			
			TransferAnimations(oDocCtrl.getCurrentPage(), oOriginalShape, oGraphic) 
		End If

		' Preserve Z index of equation
		If EditEquation Then
			oGraphic.ZOrder = oOriginalShapeZOrder
 		End If

	End If
	
	' If Writer is installed, restore initial clipboard content
	If Not (getGUIType() = 1 And sVersion = "5") Then ' Doesn't work on Windows with Libreoffice 5.x

		If ComponentInstalled( "Writer" ) Then
			TextToClipboard(sClipContent)
		End If
	
	End If

	MakeEquation = 0
	Exit Function

	' To ensure compatibility with previous TexMaths versions
	SelectionError:
		oSelection = oDocCtrl.getSelection
		Resume Next

End Function