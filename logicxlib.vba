
Option Explicit
Private oDlgMain as Variant

Sub Main
    BasicLibraries.LoadLibrary("LogicxLibrary")
    oDlgMain = LoadDialog("LogicxLibrary", "MainUi")
    oDlgMain.Execute()
End Sub
Function MakeShape(sShapeCode as String)

	print sShapeCode
	MakeShape =0
	Exit Function
End Function
Sub GenerateImage
	Dim sShapeCode as string
	sShapeCode = oDlgMain.getControl("shapeCodeTectBox").getText()
	print sShapeCode
	Dim ret as Integer
	ret = MakeShape()
	If ret = 0 Then
		oDlgMain.endExecute(sShapeCode)
	End If
End Sub