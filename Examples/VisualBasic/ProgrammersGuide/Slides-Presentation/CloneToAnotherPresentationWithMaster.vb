'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace VisualBasic.Slides
    Public Class CloneToAnotherPresentationWithMaster
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations()

            'Instantiate Presentation class to load the source presentation file

            Using srcPres As New Presentation(dataDir & "CloneToAnotherPresentationWithMaster.pptx")

                'Instantiate Presentation class for destination presentation (where slide is to be cloned)

                Using destPres As New Presentation()

                    'Instantiate ISlide from the collection of slides in source presentation along with
                    'master slide
                    Dim SourceSlide As ISlide = srcPres.Slides(0)
                    Dim SourceMaster As IMasterSlide = SourceSlide.LayoutSlide.MasterSlide

                    'Clone the desired master slide from the source presentation to the collection of masters in the
                    'destination presentation
                    Dim masters As IMasterSlideCollection = destPres.Masters
                    Dim DestMaster As IMasterSlide = SourceSlide.LayoutSlide.MasterSlide

                    'Clone the desired master slide from the source presentation to the collection of masters in the
                    'destination presentation
                    Dim iSlide As IMasterSlide = masters.AddClone(SourceMaster)

                    'Clone the desired slide from the source presentation with the desired master to the end of the
                    'collection of slides in the destination presentation
                    Dim slds As ISlideCollection = destPres.Slides
                    slds.AddClone(SourceSlide, iSlide, True)
                    'Clone the desired master slide from the source presentation to the collection of masters in the//destination presentation
                    'Save the destination presentation to disk
                    destPres.Save(dataDir & "Output_CloneToAnotherPresentationWithMaster.pptx", SaveFormat.Pptx)

                End Using
            End Using
        End Sub
    End Class
End Namespace