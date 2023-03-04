// ---------------------------------------------------------------------------
// <copyright file="MyDocument.cs" company="Tethys">
//   Copyright (C) 2023 T. Graf
// </copyright>
//
// Licensed under the Apache License, Version 2.0.
// Unless required by applicable law or agreed to in writing,
// software distributed under the License is distributed on an
// "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND,
// either express or implied.
// SPDX-License-Identifier: Apache-2.0
// ---------------------------------------------------------------------------

namespace Tethys.DocxSupport.Demo
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal class MyDocument : BasicWordSupport
    {
        /// <summary>
        /// Creates the full new document.
        /// </summary>
        /// <param name="outputFilename">The output filename.</param>
        public void CreateFullDocument(string outputFilename)
        {
            using (var doc = WordprocessingDocument.Create(outputFilename, WordprocessingDocumentType.Document))
            {
                if (doc.MainDocumentPart == null)
                {
                    BasicWordSupport.CreateNewDocument(doc);
                } // if

                if (doc.MainDocumentPart.Document == null)
                {
                    doc.MainDocumentPart.Document = new Document(new Body(new Paragraph(new Run(new Text(string.Empty)))));
                } // if

                AddUnnumberedListItem(doc.MainDocumentPart, 1, "First Level");
                AddUnnumberedListItem(doc.MainDocumentPart, 2, "Second Level");
                AddUnnumberedListItem(doc.MainDocumentPart, 3, "Third Level");

                // Add a paragraph with some text.
                var para1 = new Paragraph();
                doc.MainDocumentPart.Document.Append(para1);
                var para = doc.MainDocumentPart.Document.InsertAfter(new Paragraph(), para1);
                var run = para.AppendChild(new Run());
                run.AppendChild(new Text("Hello there!"));

                const string FatStyle = "FatStyle";
                AddStylesPartToDocument(doc);
                if (!DoesStyleExist(doc, "1"))
                {
                    AddNewStyle(doc.MainDocumentPart.StyleDefinitionsPart, "1", FatStyle);
                } // if

                para = doc.MainDocumentPart.Document.InsertAfter(new Paragraph(), para);
                run = para.AppendChild(new Run());
                run.AppendChild(new Text("Some new style"));

                ApplyStyleToParagraph(doc, "1", FatStyle, para);

                doc.MainDocumentPart.Document.Save();
            } // using
        } // CreateFullDocument()
    }
}
