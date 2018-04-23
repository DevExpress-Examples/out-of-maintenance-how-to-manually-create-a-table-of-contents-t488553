using System;
using System.Windows.Controls;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.Office;
using DevExpress.Utils;
using System.Drawing;
using DevExpress.Xpf.Ribbon;

namespace TableOfContentsExample
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : DXRibbonWindow
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private void InsertCaptions(Document document)
        {   
            #region #MarkEntries
            richEditControl1.LoadDocument("Documents//Table of Contents.docx");            
            document.BeginUpdate();

            for (int i = 0; i < document.Images.Count; i++)
            {
                DocumentImage shape = document.Images[i];
                Paragraph paragraph = document.Paragraphs.Insert(shape.Range.End);

                //Insert caption to every image in the document
                DocumentRange range = document.InsertText(paragraph.Range.Start, "Image ");

                //Mark the captions with the SEQ fields
                Field field = document.Fields.Create(range.End, "SEQ  Image \\*ARABIC");
            }
            document.Fields.Update();
            document.EndUpdate();
            #endregion #MarkEntries
        }

        private void InsertTableOfEntries()
        {
            #region #TableOfImages
            Document document = richEditControl1.Document;
            document.BeginUpdate();
            Paragraph paragraph = document.Paragraphs.Insert(document.Paragraphs[1].Range.Start);

            //Insert the table of images to the beginning of the document
            document.InsertText(paragraph.Range.Start, "Table of Images" + "\r");
            Field field = document.Fields.Create(paragraph.Range.Start, "TOC \\c Image");
            field.Update();
            
            //Update all fields to apply changes
            document.Fields.Update();
            document.EndUpdate();
            #endregion #TableOfImages
        }
        
        private void UseHeadingStyles()
        {
            #region #ApplyStyleToHeaders
            richEditControl1.LoadDocument("Documents//Table of Contents.docx");
            
            //Apply the "heading 1" style to every chapter title in the document
            for (int i = 0; i < richEditControl1.Document.Paragraphs.Count; i++)
            {
                string var = richEditControl1.Document.GetText(richEditControl1.Document.Paragraphs[i].Range);
                if (var.Contains("CHAPTER "))
                {
                    richEditControl1.Document.Paragraphs[i].Style = richEditControl1.Document.ParagraphStyles["heading 1"];
                }
            }
            #endregion #ApplyStyleToHeaders
        }
        private void InsertTableOfContentsWithStyledHeaders()
        {
            #region #HeadingStyles
            Document document = richEditControl1.Document;
            document.BeginUpdate();
            Paragraph paragraph = document.Paragraphs.Insert(document.Paragraphs[1].Range.Start);
            
            //Insert the table of contents to the beginning of the document
            document.InsertText(paragraph.Range.Start, "Table of Contents (Heading Styles) " + "\r");
            Field field = document.Fields.Create(paragraph.Range.Start, "TOC \\h ");
            field.Update();

            //Update all fields to apply changes
            document.Fields.Update();
            document.EndUpdate();     
            #endregion #HeadingStyles

        }       
        private void UseOutlineLevel()
        {
            #region #SetOutlineLevel
            richEditControl1.LoadDocument("Documents//Table of Contents.docx");
            
            //Set the outline level to every chapter title in the document
            for (int i = 0; i < richEditControl1.Document.Paragraphs.Count; i++)
            {
                string var = richEditControl1.Document.GetText(richEditControl1.Document.Paragraphs[i].Range);
                if (var.Contains("CHAPTER "))
                {
                    richEditControl1.Document.Paragraphs[i].OutlineLevel = 2;
                }
            }
            #endregion #SetOutlineLevel
        }
        private void InsertTableOfContentsUsingOutline()
        {
            #region #OutlineLevel
            Document document = richEditControl1.Document;
            document.BeginUpdate();
            
            //Insert the table of contents to the beginning of the document
            Paragraph paragraph = document.Paragraphs.Insert(document.Paragraphs[1].Range.End);
            paragraph.OutlineLevel = 0;

            document.InsertText(paragraph.Range.Start, "Table of Contents (Outline Levels)" + "\r"+"\n");
            Field field = document.Fields.Create(paragraph.Range.Start, "TOC \\u");
            field.Update();

            //Update all fields to apply changes
            document.Fields.Update();
            document.EndUpdate();
            #endregion #OutlineLevel
        }        
        
        private void UseTCField()
        {
            #region #InsertTCFields
            richEditControl1.LoadDocument("Documents//Table of Contents.docx");
            int j = 1;

            //Mark every chapter title in the document by the TC field
            for (int i = 0; i < richEditControl1.Document.Paragraphs.Count; i++)
            {
                string var = richEditControl1.Document.GetText(richEditControl1.Document.Paragraphs[i].Range);

                if (var.Contains("CHAPTER "))
                {
                    Field field = richEditControl1.Document.Fields.Create(richEditControl1.Document.Paragraphs[i].Range.Start, String.Format("TC {0} \\f bvz ", j));
                    richEditControl1.Document.Fields.Update();
                    j++;
                }

            }
            #endregion#InsertTCFields
        }

        private void InsertTCTable()
        {
            #region #TCFields
            Document document = richEditControl1.Document;
            document.BeginUpdate();

            //Insert the table of contents to the beginning of the document
            Paragraph paragraph = document.Paragraphs.Insert(document.Paragraphs[1].Range.Start);
            document.InsertText(paragraph.Range.Start, "Table of Contents (TC Fields)" + "\r");
            Field field = document.Fields.Create(paragraph.Range.Start, "TOC \\f bvz " + "\r" + "\n");
            field.Update();

            //Update all fields to apply changes
            document.Fields.Update();
            document.EndUpdate();
            #endregion #TCFields

        }
        #region #events
        private void NavBarItem_Click(object sender, EventArgs e)
        {
            InsertCaptions(richEditControl1.Document);
            InsertTableOfEntries();
        }

        private void NavBarItem_Click_1(object sender, EventArgs e)
        {
            UseHeadingStyles();
            InsertTableOfContentsWithStyledHeaders();
        }

        private void NavBarItem_Click_2(object sender, EventArgs e)
        {
            UseOutlineLevel();
            InsertTableOfContentsUsingOutline();
        }

        private void NavBarItem_Click_3(object sender, EventArgs e)
        {
            UseTCField();
            InsertTCTable();
        }
        #endregion #events
    }
}
