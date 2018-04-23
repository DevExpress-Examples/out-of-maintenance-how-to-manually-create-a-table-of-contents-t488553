Imports System
Imports System.Windows.Controls
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.Office
Imports DevExpress.Utils
Imports System.Drawing
Imports DevExpress.Xpf.Ribbon

Namespace TableOfContentsExample
    ''' <summary>
    ''' Interaction logic for MainWindow.xaml
    ''' </summary>
    Partial Public Class MainWindow
        Inherits DXRibbonWindow

        Public Sub New()
            InitializeComponent()
        End Sub
        Private Sub InsertCaptions(ByVal document As Document)
'            #Region "#MarkEntries"
            richEditControl1.LoadDocument("Documents//Table of Contents.docx")
            document.BeginUpdate()

            For i As Integer = 0 To document.Images.Count - 1
                Dim shape As DocumentImage = document.Images(i)
                Dim paragraph As Paragraph = document.Paragraphs.Insert(shape.Range.End)

                'Insert caption to every image in the document
                Dim range As DocumentRange = document.InsertText(paragraph.Range.Start, "Image ")

                'Mark the captions with the SEQ fields
                Dim field As Field = document.Fields.Create(range.End, "SEQ  Image \*ARABIC")
            Next i
            document.Fields.Update()
            document.EndUpdate()
'            #End Region ' #MarkEntries
        End Sub

        Private Sub InsertTableOfEntries()
'            #Region "#TableOfImages"
            Dim document As Document = richEditControl1.Document
            document.BeginUpdate()
            Dim paragraph As Paragraph = document.Paragraphs.Insert(document.Paragraphs(1).Range.Start)

            'Insert the table of images to the beginning of the document
            document.InsertText(paragraph.Range.Start, "Table of Images" & ControlChars.Cr)
            Dim field As Field = document.Fields.Create(paragraph.Range.Start, "TOC \c Image")
            field.Update()

            'Update all fields to apply changes
            document.Fields.Update()
            document.EndUpdate()
'            #End Region ' #TableOfImages
        End Sub


        Private Sub UseHeadingStyles()
'            #Region "#ApplyStyleToHeaders"
            richEditControl1.LoadDocument("Documents//Table of Contents.docx")

            'Apply the "heading 1" style to every chapter title in the document
            For i As Integer = 0 To richEditControl1.Document.Paragraphs.Count - 1
                Dim var As String = richEditControl1.Document.GetText(richEditControl1.Document.Paragraphs(i).Range)
                If var.Contains("CHAPTER ") Then
                    richEditControl1.Document.Paragraphs(i).Style = richEditControl1.Document.ParagraphStyles("heading 1")
                End If
            Next i
'            #End Region ' #ApplyStyleToHeaders
        End Sub
        Private Sub InsertTableOfContentsWithStyledHeaders()
'            #Region "#HeadingStyles"
            Dim document As Document = richEditControl1.Document
            document.BeginUpdate()
            Dim paragraph As Paragraph = document.Paragraphs.Insert(document.Paragraphs(1).Range.Start)

            'Insert the table of contents to the beginning of the document
            document.InsertText(paragraph.Range.Start, "Table of Contents (Heading Styles) " & ControlChars.Cr)
            Dim field As Field = document.Fields.Create(paragraph.Range.Start, "TOC \h ")
            field.Update()

            'Update all fields to apply changes
            document.Fields.Update()
            document.EndUpdate()
'            #End Region ' #HeadingStyles

        End Sub
        Private Sub UseOutlineLevel()
'            #Region "#SetOutlineLevel"
            richEditControl1.LoadDocument("Documents//Table of Contents.docx")

            'Set the outline level to every chapter title in the document
            For i As Integer = 0 To richEditControl1.Document.Paragraphs.Count - 1
                Dim var As String = richEditControl1.Document.GetText(richEditControl1.Document.Paragraphs(i).Range)
                If var.Contains("CHAPTER ") Then
                    richEditControl1.Document.Paragraphs(i).OutlineLevel = 2
                End If
            Next i
'            #End Region ' #SetOutlineLevel
        End Sub
        Private Sub InsertTableOfContentsUsingOutline()
'            #Region "#OutlineLevel"
            Dim document As Document = richEditControl1.Document
            document.BeginUpdate()

            'Insert the table of contents to the beginning of the document
            Dim paragraph As Paragraph = document.Paragraphs.Insert(document.Paragraphs(1).Range.End)
            paragraph.OutlineLevel = 0

            document.InsertText(paragraph.Range.Start, "Table of Contents (Outline Levels)" & ControlChars.Cr)
            Dim field As Field = document.Fields.Create(paragraph.Range.Start, "TOC \u ")
            field.Update()

            'Update all fields to apply changes
            document.Fields.Update()
            document.EndUpdate()
'            #End Region ' #OutlineLevel
        End Sub


        Private Sub UseTCField()
'            #Region "#InsertTCFields"
            richEditControl1.LoadDocument("Documents//Table of Contents.docx")
            Dim j As Integer = 1

            'Mark every chapter title in the document by the TC field
            For i As Integer = 0 To richEditControl1.Document.Paragraphs.Count - 1
                Dim var As String = richEditControl1.Document.GetText(richEditControl1.Document.Paragraphs(i).Range)

                If var.Contains("CHAPTER ") Then
                    Dim field As Field = richEditControl1.Document.Fields.Create(richEditControl1.Document.Paragraphs(i).Range.Start, String.Format("TC {0} \f bvz ", j))
                    richEditControl1.Document.Fields.Update()
                    j += 1
                End If

            Next i
            '#End Region '#InsertTCFields
        End Sub

        Private Sub InsertTCTable()
'            #Region "#TCFields"
            Dim document As Document = richEditControl1.Document
            document.BeginUpdate()

            'Insert the table of contents to the beginning of the document
            Dim paragraph As Paragraph = document.Paragraphs.Insert(document.Paragraphs(1).Range.Start)
            document.InsertText(paragraph.Range.Start, "Table of Contents (TC Fields)" & ControlChars.Cr)
            Dim field As Field = document.Fields.Create(paragraph.Range.Start, "TOC \f bvz " & ControlChars.Cr & ControlChars.Lf)
            field.Update()

            'Update all fields to apply changes
            document.Fields.Update()
            document.EndUpdate()
'            #End Region ' #TCFields

        End Sub
        ' #Region '"#events"
        Private Sub NavBarItem_Click(ByVal sender As Object, ByVal e As EventArgs)
            InsertCaptions(richEditControl1.Document)
            InsertTableOfEntries()
        End Sub

        Private Sub NavBarItem_Click_1(ByVal sender As Object, ByVal e As EventArgs)
            UseHeadingStyles()
            InsertTableOfContentsWithStyledHeaders()
        End Sub

        Private Sub NavBarItem_Click_2(ByVal sender As Object, ByVal e As EventArgs)
            UseOutlineLevel()
            InsertTableOfContentsUsingOutline()
        End Sub

        Private Sub NavBarItem_Click_3(ByVal sender As Object, ByVal e As EventArgs)
            UseTCField()
            InsertTCTable()
        End Sub
        '#End Region ' #events
    End Class
End Namespace
