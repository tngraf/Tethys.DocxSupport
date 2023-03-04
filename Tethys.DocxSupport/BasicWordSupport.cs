// ---------------------------------------------------------------------------
// <copyright file="BasicWordSupport.cs" company="Tethys">
//   Copyright (C) 2016-2023 T. Graf
// </copyright>
//
// Licensed under the Apache License, Version 2.0.
// Unless required by applicable law or agreed to in writing,
// software distributed under the License is distributed on an
// "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND,
// either express or implied.
// SPDX-License-Identifier: Apache-2.0
// ---------------------------------------------------------------------------

/*****************************************************************************
 * Required NuGet Packages
 * -----------------------
 * - DocumentFormat.OpenXml 2.7.2
 ****************************************************************************/

/*****************************************************************************
* Things that are **NOT** possible:
* -----------------------
* - It is not possible to generate a table of contents (TOC)
* - It is not possible to run a macro
*
****************************************************************************/

namespace Tethys.DocxSupport
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Reflection;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.CustomProperties;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.VariantTypes;
    using DocumentFormat.OpenXml.Wordprocessing;

    using Tethys.Logging;

    /// <summary>
    /// Generator methods for Word documents.
    /// </summary>
    public class BasicWordSupport
    {
        #region PRIVATE PROPERTIES
        /// <summary>
        /// The logger for this class.
        /// </summary>
        private static readonly ILog Log = LogManager.GetLogger(typeof(BasicWordSupport));
        #endregion // PRIVATE PROPERTIES

        //// ---------------------------------------------------------------------

        #region PUBLIC PROPERTIES
        /// <summary>
        /// Property types for <c>docx</c> custom properties.
        /// </summary>
        public enum PropertyTypes
        {
            /// <summary>
            /// A yes/no value.
            /// </summary>
            YesNo,

            /// <summary>
            /// A text value.
            /// </summary>
            Text,

            /// <summary>
            /// A date/time value.
            /// </summary>
            DateTime,

            /// <summary>
            /// An integer value.
            /// </summary>
            NumberInteger,

            /// <summary>
            /// A double value.
            /// </summary>
            NumberDouble,
        } // PropertyTypes
        #endregion // PUBLIC PROPERTIES

        //// ---------------------------------------------------------------------

        #region PUBLIC METHODS
        /// <summary>
        /// Opens the document in word.
        /// </summary>
        /// <param name="filename">The fileName.</param>
        public static void OpenDocumentInWord(string filename)
        {
            Log.InfoFormat("Opening Microsoft Word for file '{0}'", filename);

            try
            {
                var process = new Process();
                process.StartInfo.UseShellExecute = true;
                process.StartInfo.FileName = "WINWORD.EXE";
                process.StartInfo.Arguments = $"\"{filename}\"";
                process.Start();
            }
            catch (Exception ex)
            {
                Log.Error("Error opening Microsoft Word", ex);
            } // catch
        } // OpenDocumentInWord()
        #endregion // PUBLIC METHODS

        //// ---------------------------------------------------------------------

        #region PROTECTED METHODS
        /// <summary>
        /// Applies the style to paragraph.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="styleId">The style id.</param>
        /// <param name="styleName">The styleName.</param>
        /// <param name="p">The p.</param>
        public static void ApplyStyleToParagraph(
            WordprocessingDocument doc,
            string styleId,
            string styleName,
            Paragraph p)
        {
            // If the paragraph has no ParagraphProperties object, create one.
            if (!p.Elements<ParagraphProperties>().Any())
            {
                p.PrependChild(new ParagraphProperties());
            } // if

            // Get the paragraph properties element of the paragraph.
            var pPr = p.Elements<ParagraphProperties>().First();
            pPr.ParagraphStyleId = new ParagraphStyleId { Val = styleId };
        } // ApplyStyleToParagraph()

        /// <summary>
        /// Adds the styles part to the document.
        /// </summary>
        /// <param name="doc">The document.</param>
        public static void AddStylesPartToDocument(WordprocessingDocument doc)
        {
            // If the Styles part does not exist, add it and then add the style.
            if (doc.MainDocumentPart.StyleDefinitionsPart != null)
            {
                return;
            } // if

            var part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            var root = new Styles();
            root.Save(part);
        } // AddStylesPartToDocument()

        /// <summary>
        /// Checks whether the given style exist in the document.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="styleId">The style id.</param>
        /// <returns>true if the style exists; otherwise false.</returns>
        public static bool DoesStyleExist(WordprocessingDocument doc, string styleId)
        {
            // Get access to the Styles element for this document.
            var s = doc.MainDocumentPart.StyleDefinitionsPart.Styles;

            // Check that there are styles and how many.
            var n = s.Elements<Style>().Count();
            if (n == 0)
            {
                return false;
            } // if

            // Look for a match on styleId.
            var style = s.Elements<Style>()
                .FirstOrDefault(st => (st.StyleId == styleId) && (st.Type == StyleValues.Paragraph));
            return style != null;
        } // DoesStyleExist()

        /// <summary>
        /// Create a new style with the specified styleId and styleName and add it to
        /// the specified style definitions part.
        /// </summary>
        /// <param name="styleDefinitionsPart">The style definitions part.</param>
        /// <param name="styleId">The style id.</param>
        /// <param name="styleName">The style name.</param>
        public static void AddNewStyle(
            StyleDefinitionsPart styleDefinitionsPart,
            string styleId,
            string styleName)
        {
            // Get access to the root element of the styles part.
            var styles = styleDefinitionsPart.Styles;

            // Create a new paragraph style and specify some of the properties.
            var style = new Style
                            {
                                Type = StyleValues.Paragraph,
                                StyleId = styleId,
                                CustomStyle = true,
                            };
            var styleName1 = new StyleName { Val = styleName };
            var basedOn1 = new BasedOn { Val = "Normal" };
            var nextParagraphStyle1 = new NextParagraphStyle { Val = "Normal" };
            style.Append(styleName1);
            style.Append(basedOn1);
            style.Append(nextParagraphStyle1);

            // Create the StyleRunProperties object and specify some of the run properties.
            var styleRunProperties1 = new StyleRunProperties();
            var bold1 = new Bold();
            var color1 = new Color { ThemeColor = ThemeColorValues.Accent2 };
            var font1 = new RunFonts { Ascii = "Lucida Console" };
            var italic1 = new Italic();

            // Specify a 12 point size.
            var fontSize1 = new FontSize { Val = "24" };
            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(font1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(italic1);

            // Add the run properties to the style.
            style.Append(styleRunProperties1);

            // Add the style to the styles part.
            styles.Append(style);
        } // AddNewStyle()

        /// <summary>
        /// Sets the font properties of the given <see cref="Run" />.
        /// </summary>
        /// <param name="run">The run.</param>
        /// <param name="fontName">Name of the font.</param>
        public static void SetRunFontProperties(Run run, string fontName)
        {
            var rPr = new RunProperties(
                new RunFonts
                    {
                        Ascii = fontName,
                    });

            run.PrependChild(rPr);
        } // SetRunFontProperties()

        /// <summary>
        /// Sets the font properties of the given <see cref="Run" />.
        /// </summary>
        /// <param name="run">The run.</param>
        /// <param name="fontName">Name of the font.</param>
        /// <param name="fontSize">Size of the font.</param>
        /// <remarks>
        /// Note that the site needs to be specified in half-points (1/144 of an inch).
        /// 40 half-points are a font size of 20.
        /// </remarks>
        public static void SetRunFontProperties(Run run, string fontName, string fontSize)
        {
            var rPr = new RunProperties(
                new RunFonts
                    {
                        Ascii = fontName,
                    });

            var fs = new FontSize();
            fs.Val = fontSize;
            rPr.AppendChild(fs);

            run.PrependChild(rPr);
        } // SetRunFontProperties()

        /// <summary>
        /// Sets a custom property.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="propertyName">Name of the property.</param>
        /// <param name="propertyValue">The property value.</param>
        /// <param name="propertyType">Type of the property.</param>
        /// <returns>The previous value of this property.</returns>
        public static string SetCustomProperty(
            WordprocessingDocument doc,
            string propertyName,
            object propertyValue,
            PropertyTypes propertyType)
        {
            // see https://docs.microsoft.com/de-de/office/open-xml/how-to-set-a-custom-property-in-a-word-processing-document
            string returnValue = null;

            var newProp = new CustomDocumentProperty();
            var propSet = false;

            // Calculate the correct type.
            switch (propertyType)
            {
                case PropertyTypes.DateTime:
                    // Be sure you were passed a real date,
                    // and if so, format in the correct way.
                    // The date/time value passed in should
                    // represent a UTC date/time.
                    if (propertyValue is DateTime)
                    {
                        newProp.VTFileTime =
                            new VTFileTime($"{Convert.ToDateTime(propertyValue):s}Z");
                        propSet = true;
                    } // if

                    break;
                case PropertyTypes.NumberInteger:
                    if (propertyValue is int)
                    {
                        newProp.VTInt32 = new VTInt32(propertyValue.ToString());
                        propSet = true;
                    } // if

                    break;
                case PropertyTypes.NumberDouble:
                    if (propertyValue is double)
                    {
                        newProp.VTFloat = new VTFloat(propertyValue.ToString());
                        propSet = true;
                    } // if

                    break;

                case PropertyTypes.Text:
                    newProp.VTLPWSTR = new VTLPWSTR(propertyValue.ToString());
                    propSet = true;

                    break;
                case PropertyTypes.YesNo:
                    if (propertyValue is bool)
                    {
                        // Must be lowercase.
                        newProp.VTBool = new VTBool(
                            Convert.ToBoolean(propertyValue).ToString().ToLower());
                        propSet = true;
                    } // if

                    break;
            } // switch

            if (!propSet)
            {
                // If the code was not able to convert the
                // property to a valid value, throw an exception.
                throw new InvalidDataException("propertyValue");
            } // if

            // Now that you have handled the parameters, start
            // working on the document.
            newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
            newProp.Name = propertyName;

            var customProps = doc.CustomFilePropertiesPart;
            if (customProps == null)
            {
                // No custom properties? Add the part, and the
                // collection of properties now.
                customProps = doc.AddCustomFilePropertiesPart();
                customProps.Properties = new Properties();
            } // if

            var props = customProps.Properties;
            if (props != null)
            {
                // This will trigger an exception if the property's Name
                // property is null, but if that happens, the property is damaged,
                // and probably should raise an exception.
                var prop = props.FirstOrDefault(p => ((CustomDocumentProperty)p).Name.Value
                                                     == propertyName);

                // Does the property exist? If so, get the return value,
                // and then delete the property.
                if (prop != null)
                {
                    returnValue = prop.InnerText;
                    prop.Remove();
                } // if

                // Append the new property, and
                // fix up all the property ID values.
                // The PropertyId value must start at 2.
                props.AppendChild(newProp);
                var pid = 2;
                foreach (var openXmlElement in props)
                {
                    var item = (CustomDocumentProperty)openXmlElement;
                    item.PropertyId = pid++;
                } // foreach

                props.Save();
            } // if

            return returnValue;
        } // SetCustomProperty()

        /// <summary>
        /// Validates the word document.
        /// </summary>
        /// <param name="filepath">The file path.</param>
        /// <returns>The number of errors found.</returns>
        public static int ValidateWordDocument(string filepath)
        {
            // https://docs.microsoft.com/de-de/office/open-xml/how-to-validate-a-word-processing-document
            int count;
            using (var doc = WordprocessingDocument.Open(filepath, true))
            {
                count = ValidateWordDocument(doc);
            } // using

            return count;
        } // ValidateWordDocument()

        /// <summary>
        /// Validates the word document.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <returns>
        /// The number of errors found.
        /// </returns>
        public static int ValidateWordDocument(WordprocessingDocument doc)
        {
            var count = 0;
            try
            {
                Log.Debug("Validating document...");
                var validator = new OpenXmlValidator();
                foreach (var error in validator.Validate(doc))
                {
                    count++;
                    Log.Error("Error " + count);
                    Log.Error("Description: " + error.Description);
                    Log.Error("ErrorType: " + error.ErrorType);
                    Log.Error("Node: " + error.Node);
                    Log.Error("Path: " + error.Path.XPath);
                    Log.Error("Part: " + error.Part.Uri);
                    Log.Error("-------------------------------------------");
                } // foreach

                if (count > 0)
                {
                    Log.Error($"Total issue count={count}");
                } // if
            }
            catch (Exception ex)
            {
                Log.Error("Error validating document: " + ex.Message);
            } // catch

            return count;
        } // ValidateWordDocument()

        /// <summary>
        /// Finds the given text in the given document.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="searchText">The search text.</param>
        /// <param name="perfectMatch">if set to <c>true</c> look for a perfect match.</param>
        /// <returns>The <see cref="Paragraph"/> or null.</returns>
        protected static Paragraph FindText(WordprocessingDocument doc, string searchText, bool perfectMatch)
        {
            var body = doc.MainDocumentPart.Document.Body;
            var paras = body.Elements<Paragraph>();

            foreach (var para in paras)
            {
                foreach (var run in para.Elements<Run>())
                {
                    foreach (var text in run.Elements<Text>())
                    {
                        if (perfectMatch)
                        {
                            if (text.Text.Equals(searchText, StringComparison.CurrentCulture))
                            {
                                return para;
                            } // if
                        }
                        else
                        {
                            if (text.Text.Contains(searchText))
                            {
                                return para;
                            } // if
                        } // if
                    } // foreach
                } // foreach
            } // foreach

            return null;
        } // FindText()

        /// <summary>
        /// Finds the given text in the given table.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <param name="searchText">The search text.</param>
        /// <param name="perfectMatch">if set to <c>true</c> look for a perfect match.</param>
        /// <returns>The <see cref="Paragraph"/> or null.</returns>
        protected static Paragraph FindTextInTable(Table table, string searchText, bool perfectMatch)
        {
            if (table == null)
            {
                throw new ArgumentNullException();
            } // if

            var rows = table.Elements<TableRow>();
            foreach (var row in rows)
            {
                foreach (var cell in row.Elements<TableCell>())
                {
                    var paras = cell.Elements<Paragraph>();

                    foreach (var para in paras)
                    {
                        foreach (var run in para.Elements<Run>())
                        {
                            foreach (var text in run.Elements<Text>())
                            {
                                if (perfectMatch)
                                {
                                    if (text.Text.Equals(searchText, StringComparison.CurrentCulture))
                                    {
                                        return para;
                                    } // if
                                }
                                else
                                {
                                    if (text.Text.Contains(searchText))
                                    {
                                        return para;
                                    } // if
                                } // if
                            } // foreach
                        } // foreach
                    } // foreach
                } // foreach
            } // foreach

            return null;
        } // FindTextInTable()

        /// <summary>
        /// Highlights the text in the given paragraph.
        /// </summary>
        /// <param name="para">The paragraph.</param>
        protected static void HighlightParagraph(Paragraph para)
        {
            var run = para.GetFirstChild<Run>();
            if (run == null)
            {
                return;
            } // if

            // remove old data
            var oldProperties = run.GetFirstChild<RunProperties>();
            oldProperties?.Remove();

            var text = run.GetFirstChild<Text>();
            if (text == null)
            {
                return;
            } // if

            var plaintext = text.Text;
            text.Remove();
            run.Remove();

            // create new run properties
            var h = new Highlight();
            h.Val = HighlightColorValues.Yellow;

            var rp = new RunProperties();
            rp.AppendChild(h);

            // add new run with new properties
            para.Append(new Run(rp, new Text(plaintext)));
        } // HighlightParagraph()

        /// <summary>
        /// Creates a new document.
        /// </summary>
        /// <param name="wordDoc">The word document.</param>
        protected static void CreateNewDocument(WordprocessingDocument wordDoc)
        {
            wordDoc.AddMainDocumentPart();

            wordDoc.MainDocumentPart.Document = new Document();
            /*Body body = */
            wordDoc.MainDocumentPart.Document.AppendChild(new Body());

            // new Paragraph
            // new Run
            // new Text
        } // CreateNewDocument()

        /// <summary>
        /// Adds the yes no checkboxes.
        /// </summary>
        /// <param name="mdp">The MDP.</param>
        /// <param name="para">The para.</param>
        protected static void AddYesNoCheckboxes(MainDocumentPart mdp, OpenXmlElement para)
        {
            AddCheckBoxToParagraph(para, "yes   ", true);
            AddCheckBoxToParagraph(para, "no", false);

            mdp.Document.Body.Append(para);
        } // AddYesNoCheckboxes()

        /// <summary>
        /// Adds a new check box to the given paragraph.
        /// </summary>
        /// <param name="para">The para.</param>
        /// <param name="text">The text.</param>
        /// <param name="value">if set to <c>true</c> [value].</param>
        protected static void AddCheckBoxToParagraph(OpenXmlElement para, string text, bool value)
        {
            var fc = new FieldChar(
                new FormFieldData(
                    new FormFieldName { Val = "Check1" },
                    new Enabled(),
                    new CalculateOnExit { Val = OnOffValue.FromBoolean(false) },
                    new CheckBox(
                        new AutomaticallySizeFormField(),
                        new DefaultCheckBoxFormFieldState { Val = OnOffValue.FromBoolean(value) })));
            fc.FieldCharType = FieldCharValues.Begin;

            var tx = new Text(text);
            tx.Space = SpaceProcessingModeValues.Preserve;

            para.Append(new Run(fc));
            para.Append(new BookmarkStart { Name = "Check1", Id = "0" });
            para.Append(
                new Run(
                    new FieldCode(" FORMCHECKBOX ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(
                    new FieldChar { FieldCharType = FieldCharValues.End }),
                new BookmarkEnd { Id = "0" },
                new Run(tx));
            ////{ RsidParagraphAddition = "003865D4" };
        } // AddCheckBoxToParagraph()

        /// <summary>
        /// Adds the unnumbered list item.
        /// </summary>
        /// <param name="mdp">The MDP.</param>
        /// <param name="level">The level.</param>
        /// <param name="text">The text.</param>
        /// <returns>The newly added <see cref="Paragraph"/>.</returns>
        protected static Paragraph AddUnnumberedListItem(MainDocumentPart mdp, int level, string text)
        {
            var para = new Paragraph(
                new ParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference { Val = level },
                        new NumberingId { Val = 4 })),
                new Run(new RunProperties(), new Text(text)));
            mdp.Document.Body.Append(para);
            return para;
        } // AddUnnumberedListItem()

        /// <summary>
        /// Inserts an unnumbered list item.
        /// </summary>
        /// <param name="mdp">The MDP.</param>
        /// <param name="level">The level.</param>
        /// <param name="text">The text.</param>
        /// <param name="insertAfter">The position after which the new data shall get inserted.</param>
        /// <returns>
        /// The newly added <see cref="Paragraph" />.
        /// </returns>
        protected static Paragraph InsertUnnumberedListItem(MainDocumentPart mdp, int level, string text, Paragraph insertAfter)
        {
            var para = new Paragraph(
                new ParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference { Val = level },
                        new NumberingId { Val = 4 })),
                new Run(new RunProperties(), new Text(text)));
            return mdp.Document.Body.InsertAfter(para, insertAfter);
        } // InsertUnnumberedListItem()

        /// <summary>
        /// Adds the given text in italic.
        /// </summary>
        /// <param name="mdp">The MDP.</param>
        /// <param name="text">The text.</param>
        /// <returns>The newly added <see cref="Paragraph"/>.</returns>
        protected static Paragraph AddItalicText(MainDocumentPart mdp, string text)
        {
            var run = new Run();
            var runProperties = run.AppendChild(new RunProperties());
            var italic = new Italic();
            italic.Val = OnOffValue.FromBoolean(true);
            runProperties.AppendChild(italic);
            run.AppendChild(new Text(text));

            var para = new Paragraph(run);
            mdp.Document.Body.Append(para);
            return para;
        } // AddItalicText()

        /// <summary>
        /// Adds the header with the given text based on the given style.
        /// </summary>
        /// <param name="mdp">The MDP.</param>
        /// <param name="text">The text.</param>
        /// <param name="style">The style.</param>
        /// <returns>The newly added <see cref="Paragraph"/>.</returns>
        protected static Paragraph AddTemplateHeader(MainDocumentPart mdp, string text, string style)
        {
            // This methods requires that the style is ALREADY defined in
            // the word document (template)
            // ==> see document (unzipped), file word/styles.xml,
            // look for 'w:styleId'
            var para = mdp.Document.Body.AppendChild(new Paragraph());
            var run = para.AppendChild(new Run());
            run.AppendChild(new Text(text));
            para.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId { Val = style });
            return para;
        } // AddTemplateHeader()

        /// <summary>
        /// Inserts the header with the given text based on the given style.
        /// </summary>
        /// <param name="mdp">The MDP.</param>
        /// <param name="text">The text.</param>
        /// <param name="style">The style.</param>
        /// <param name="insertAfter">The position after which the new data shall get inserted.</param>
        /// <returns>The newly inserted <see cref="Paragraph"/>.</returns>
        protected static Paragraph InsertTemplateHeader(MainDocumentPart mdp, string text, string style, Paragraph insertAfter)
        {
            // This methods requires that the style is ALREADY defined in
            // the word document (template)
            // ==> see document (unzipped), file word/styles.xml,
            // look for 'w:styleId'
            var para = mdp.Document.Body.InsertAfter(new Paragraph(), insertAfter);
            var run = para.AppendChild(new Run());
            run.AppendChild(new Text(text));
            para.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId { Val = style });
            return para;
        } // InsertTemplateHeader()

        /// <summary>
        /// Adds the table.
        /// </summary>
        /// <param name="mdp">The main document part.</param>
        /// <param name="data">The data.</param>
        /// <returns>A <see cref="Table"/> object.</returns>
        protected static Table AddTable(MainDocumentPart mdp, IReadOnlyList<Tuple<string, string>> data)
        {
            var table = new Table();

            var props = new TableProperties(
                new TableBorders(
                    new TopBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4,
                    },
                    new BottomBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4,
                    },
                    new LeftBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4,
                    },
                    new RightBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4,
                    },
                    new InsideHorizontalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4,
                    },
                    new InsideVerticalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4,
                    }));

            table.AppendChild(props);

            for (var i = 0; i < data.Count; i++)
            {
                var tr = new TableRow();
                table.Append(tr);

                var tc = new TableCell();
                tr.Append(tc);
                tc.Append((i == 0) ? GetBoldParagraph(data[i].Item1) : GetParagraph(data[i].Item1));

                tc = new TableCell();
                tr.Append(tc);
                tc.Append((i == 0) ? GetBoldParagraph(data[i].Item2) : GetParagraph(data[i].Item2));
            } // for (i)

            mdp.Document.Body.Append(table);
            return table;
        } // AddTable()

        /// <summary>
        /// Gets a new paragraph for the given text.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <returns>A <see cref="Paragraph"/>.</returns>
        protected static Paragraph GetParagraph(string text)
        {
            var para = new Paragraph(
                new Run(new RunProperties(), new Text(text)));
            return para;
        } // GetParagraph()

        /// <summary>
        /// Inserts the paragraph.
        /// </summary>
        /// <param name="mdp">The main document part.</param>
        /// <param name="insertAfter">The position after which the new data shall get inserted.</param>
        /// <param name="text">The text.</param>
        /// <returns>The last added paragraph.</returns>
        protected static Paragraph InsertParagraph(MainDocumentPart mdp, Paragraph insertAfter, string text)
        {
            var para = mdp.Document.Body.InsertAfter(new Paragraph(), insertAfter);
            var run = para.AppendChild(new Run());
            run.AppendChild(new Text(text));
            return para;
        } // InsertParagraph()

        /// <summary>
        /// Gets a new bold paragraph for the given text.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <returns>A <see cref="Paragraph"/>.</returns>
        protected static Paragraph GetBoldParagraph(string text)
        {
            var run = new Run();
            var runProperties = run.AppendChild(new RunProperties());
            var bold = new Bold();
            bold.Val = OnOffValue.FromBoolean(true);
            runProperties.AppendChild(bold);
            run.AppendChild(new Text(text));

            var para = new Paragraph(run);
            return para;
        } // GetBoldParagraph()

        /// <summary>
        /// Adds the given text as new paragraph.
        /// </summary>
        /// <param name="body">The body.</param>
        /// <param name="text">The text.</param>
        /// <returns>A <see cref="Paragraph"/>.</returns>
        protected static Paragraph AddParagraph(OpenXmlElement body, string text)
        {
            var para = GetParagraph(text);
            body.Append(para);
            return para;
        } // AddParagraph()

        /// <summary>
        /// Adds a line break.
        /// </summary>
        /// <param name="para">The paragraph.</param>
        /// <returns>A <see cref="Paragraph"/>.</returns>
        protected static Paragraph AddLineBreak(Paragraph para)
        {
            para.Append(new Run(new Break { Type = BreakValues.Page }));

            return para;
        } // AddLineBreak()

        /// <summary>
        /// Copies the template file to the target file.
        /// </summary>
        /// <param name="templateName">Name of the template.</param>
        /// <param name="fileName">Name of the file.</param>
        protected static void CopyTemplateToTarget(string templateName, string fileName)
        {
            var location = Assembly.GetExecutingAssembly().Location;
            var fi = new FileInfo(location);
            var folder = fi.DirectoryName;
            if (folder == null)
            {
                Log.Error("Error getting template path!");
                return;
            } // if

            var source = Path.Combine(folder, templateName);
            File.Copy(source, fileName, true);
        } // CopyTemplateToTarget()

        /// <summary>
        /// Gets the header style by numeric level.
        /// </summary>
        /// <param name="level">The level.</param>
        /// <returns>The correspnsding style name.</returns>
        protected static string GetHeaderStyleByLevel(int level)
        {
            // ReSharper disable StringLiteralTypo
            switch (level)
            {
                case 1:
                    return "berschrift1";
                case 2:
                    return "berschrift2";
                case 3:
                    return "berschrift3";
                case 4:
                    return "berschrift4";
                case 5:
                    return "berschrift5";
                default:
                    Log.Warn($"Unknown header style requested: {level}");
                    return "berschrift1";
            } // switch

            // ReSharper enable StringLiteralTypo
        } // GetHeaderStyleByLevel()
        #endregion // PROTECTED METHODS
    } // BasicWordSupport
}
