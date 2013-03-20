'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Pptx

Namespace SavingAPresentationEx
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")



			' 1.
			' Save presentation to file.

			'Instantiate a Presentation object that represents a PPT file
			Dim pres1 As New PresentationEx()

			'...do some work here...
			ProcessPresentation(pres1)

			'Save your presentation to a file
			pres1.Write(dataDir & "toFile.pptx")



			' 2. 
			' Save your presentation to a stream

			'Instantiate a Presentation object that represents a PPT file
			Dim pres2 As New PresentationEx()

			'...do some work here...
			ProcessPresentation(pres2)

			'Accessing the output stream of Http Response
			Dim st As System.IO.Stream = New FileStream(dataDir & "toStream.pptx", FileMode.OpenOrCreate)

			'Saving the presentation to the output stream of Http Response
			pres2.Write(st)

			' Close the stream.
			st.Close()



			' 3.
			' Saving a presentation with password protection.

			'Instantiate a Presentation object that represents a PPT file
			Dim pres3 As New PresentationEx()

			'...do some work here...
			ProcessPresentation(pres3)

			'Setting Password
			pres3.Encrypt("test")

			'Save your presentation to a file
			pres3.Write(dataDir & "passwordProtected.pptx")



			' 4.
			' Save password protected Presentation with Read Access to Document Properties

			'Instantiate a Presentation object that represents a PPT file
			Dim pres4 As New PresentationEx()

			'...do some work here...
			ProcessPresentation(pres4)

			'Setting access to document properties in password protected mode
			pres4.EncryptDocumentProperties = False

			'Setting Password
			pres4.Encrypt("test")

			'Save your presentation to a file
			pres4.Write(dataDir & "passwordProtectedReadOnlyProperties.pptx")



			' 5.
			' Save a read only presentation.

			'Instantiate a Presentation object that represents a PPT file
			Dim pres5 As New PresentationEx()

			'...do some work here...
			ProcessPresentation(pres5)

			'Setting Write protection Password
			pres5.SetWriteProtection("test")

			'Save your presentation to a file
			pres5.Write(dataDir & "readOnlyPresentation.pptx")



			' 6.
			' Removing Write Protection from a Presentation

			'Opening the presentation file
			Dim pres6 As New PresentationEx(dataDir & "readOnlyPresentation.pptx")

			'Checking if presentation is write protected
			If pres6.IsWriteProtected Then
				'Removing Write protection	
				pres6.RemoveWriteProtection()
			End If

			'Saving presentation
			pres6.Write(dataDir & "writeProtectionRemoved.pptx")
		End Sub

		Public Shared Sub ProcessPresentation(ByVal pres As PresentationEx)
			'Get the first slide
			Dim sld As SlideEx = pres.Slides(0)

			'Add an AutoShape of Rectangle type
			Dim idx As Integer = sld.Shapes.AddAutoShape(ShapeTypeEx.Rectangle, 150, 75, 150, 50)
			Dim ashp As AutoShapeEx = CType(sld.Shapes(idx), AutoShapeEx)

			'Add TextFrame to the Rectangle
			ashp.AddTextFrame("Aspose")

			'Change the text color to Black (which is White by default)
			ashp.TextFrame.Paragraphs(0).Portions(0).FillFormat.FillType = FillTypeEx.Solid
			ashp.TextFrame.Paragraphs(0).Portions(0).FillFormat.SolidFillColor.Color = Color.Black

			'Change the line color of the rectangle to White
			ashp.ShapeStyle.LineColor.Color = System.Drawing.Color.White

			'Remove any fill formatting in the shape
			ashp.FillFormat.FillType = FillTypeEx.NoFill
		End Sub
	End Class
End Namespace