﻿// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Experimental;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public class SlideSource
    {
        public PmlDocument PmlDocument { get; set; }
        public int Start { get; set; }
        public int Count { get; set; }
        public bool KeepMaster { get; set; }

        public SlideSource(PmlDocument source, bool keepMaster)
        {
            PmlDocument = source;
            Start = 0;
            Count = int.MaxValue;
            KeepMaster = keepMaster;
        }

        public SlideSource(string fileName, bool keepMaster)
        {
            PmlDocument = new PmlDocument(fileName);
            Start = 0;
            Count = int.MaxValue;
            KeepMaster = keepMaster;
        }

        public SlideSource(PmlDocument source, int start, bool keepMaster)
        {
            PmlDocument = source;
            Start = start;
            Count = int.MaxValue;
            KeepMaster = keepMaster;
        }

        public SlideSource(string fileName, int start, bool keepMaster)
        {
            PmlDocument = new PmlDocument(fileName);
            Start = start;
            Count = int.MaxValue;
            KeepMaster = keepMaster;
        }

        public SlideSource(PmlDocument source, int start, int count, bool keepMaster)
        {
            PmlDocument = source;
            Start = start;
            Count = count;
            KeepMaster = keepMaster;
        }

        public SlideSource(string fileName, int start, int count, bool keepMaster)
        {
            PmlDocument = new PmlDocument(fileName);
            Start = start;
            Count = count;
            KeepMaster = keepMaster;
        }
    }

    public static class PresentationBuilder
    {
        public static void BuildPresentation(List<SlideSource> sources, string fileName)
        {
            using OpenXmlMemoryStreamDocument streamDoc = OpenXmlMemoryStreamDocument.CreatePresentationDocument();
            using (PresentationDocument output = streamDoc.GetPresentationDocument())
            {
                BuildPresentation(sources, output);
            }
            streamDoc.GetModifiedDocument().SaveAs(fileName);
        }

        public static PmlDocument BuildPresentation(List<SlideSource> sources)
        {
            using OpenXmlMemoryStreamDocument streamDoc = OpenXmlMemoryStreamDocument.CreatePresentationDocument();
            using (PresentationDocument output = streamDoc.GetPresentationDocument())
            {
                BuildPresentation(sources, output);
            }
            return streamDoc.GetModifiedPmlDocument();
        }

        private static void BuildPresentation(List<SlideSource> sources, PresentationDocument output)
        {
            _relationshipMarkup ??= new Dictionary<XName, XName[]>()
                {
                    { A.audioFile,        new [] { R.link }},
                    { A.videoFile,        new [] { R.link }},
                    { A.quickTimeFile,    new [] { R.link }},
                    { A.wavAudioFile,     new [] { R.embed }},
                    { A.blip,             new [] { R.embed, R.link }},
                    { A.hlinkClick,       new [] { R.id }},
                    { A.hlinkMouseOver,   new [] { R.id }},
                    { A.hlinkHover,       new [] { R.id }},
                    { A.relIds,           new [] { R.cs, R.dm, R.lo, R.qs }},
                    { C.chart,            new [] { R.id }},
                    { C.externalData,     new [] { R.id }},
                    { C.userShapes,       new [] { R.id }},
                    { DGM.relIds,         new [] { R.cs, R.dm, R.lo, R.qs }},
                    { A14.imgLayer,       new [] { R.embed }},
                    { P14.media,          new [] { R.embed, R.link }},
                    { P.oleObj,           new [] { R.id }},
                    { P.externalData,     new [] { R.id }},
                    { P.control,          new [] { R.id }},
                    { P.snd,              new [] { R.embed }},
                    { P.sndTgt,           new [] { R.embed }},
                    { PAV.srcMedia,       new [] { R.embed, R.link }},
                    { P.contentPart,      new [] { R.id }},
                    { VML.fill,           new [] { R.id }},
                    { VML.imagedata,      new [] { R.href, R.id, R.pict, O.relid }},
                    { VML.stroke,         new [] { R.id }},
                    { WNE.toolbarData,    new [] { R.id }},
                    { Plegacy.textdata,   new [] { XName.Get("id") }},
                };

            List<ImageData> images = new List<ImageData>();
            List<MediaData> mediaList = new List<MediaData>();
            XDocument mainPart = output.PresentationPart.GetXDocument();
            mainPart.Declaration.Standalone = "yes";
            mainPart.Declaration.Encoding = "UTF-8";
            output.PresentationPart.PutXDocument();

            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(sources[0].PmlDocument))
            using (PresentationDocument doc = streamDoc.GetPresentationDocument())
            {
                CopyStartingParts(doc, output);
            }

            int sourceNum = 0;
            SlideMasterPart currentMasterPart = null;
            foreach (SlideSource source in sources)
            {
                using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(source.PmlDocument))
                using (PresentationDocument doc = streamDoc.GetPresentationDocument())
                {
                    try
                    {
                        if (sourceNum == 0)
                            CopyPresentationParts(doc, output, images, mediaList);
                        currentMasterPart = AppendSlides(doc, output, source.Start, source.Count, source.KeepMaster, images, currentMasterPart, mediaList);
                    }
                    catch (PresentationBuilderInternalException dbie)
                    {
                        if (dbie.Message.Contains("{0}"))
                            throw new PresentationBuilderException(string.Format(dbie.Message, sourceNum));
                        else
                            throw;
                    }
                }
                sourceNum++;
            }
            foreach (var part in output.GetAllParts())
            {
                if (part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slide+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.theme+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.chart+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml" ||
                    part.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml" ||
                    part.ContentType == "application/vnd.ms-office.drawingml.diagramDrawing+xml")
                {
                    XDocument xd = part.GetXDocument();
                    xd.Descendants().Attributes("smtClean").Remove();
                    part.PutXDocument();
                }
                else if (part.Annotation<XDocument>() != null)
                    part.PutXDocument();
            }
        }

        private static void CopyStartingParts(PresentationDocument sourceDocument, PresentationDocument newDocument)
        {
            // A Core File Properties part does not have implicit or explicit relationships to other parts.
            var corePart = sourceDocument.CoreFilePropertiesPart;
            if (corePart != null && corePart.GetXDocument().Root != null)
            {
                newDocument.AddCoreFilePropertiesPart();
                XDocument newXDoc = newDocument.CoreFilePropertiesPart.GetXDocument();
                newXDoc.Declaration.Standalone = "yes";
                newXDoc.Declaration.Encoding = "UTF-8";
                XDocument sourceXDoc = corePart.GetXDocument();
                newXDoc.Add(sourceXDoc.Root);
            }

            // An application attributes part does not have implicit or explicit relationships to other parts.
            var extPart = sourceDocument.ExtendedFilePropertiesPart;
            if (extPart != null)
            {
                newDocument.AddExtendedFilePropertiesPart();
                XDocument newXDoc = newDocument.ExtendedFilePropertiesPart.GetXDocument();
                newXDoc.Declaration.Standalone = "yes";
                newXDoc.Declaration.Encoding = "UTF-8";
                newXDoc.Add(extPart.GetXDocument().Root);
            }

            // An custom file properties part does not have implicit or explicit relationships to other parts.
            var customPart = sourceDocument.CustomFilePropertiesPart;
            if (customPart != null)
            {
                newDocument.AddCustomFilePropertiesPart();
                XDocument newXDoc = newDocument.CustomFilePropertiesPart.GetXDocument();
                newXDoc.Declaration.Standalone = "yes";
                newXDoc.Declaration.Encoding = "UTF-8";
                newXDoc.Add(customPart.GetXDocument().Root);
            }
        }

#if false
            // TODO need to handle the following

            { P.custShowLst, 80 },
            { P.photoAlbum, 90 },
            { P.custDataLst, 100 },
            { P.kinsoku, 120 },
            { P.modifyVerifier, 150 },
#endif

        // Copy handout master, notes master, presentation properties and view properties, if they exist
        private static void CopyPresentationParts(PresentationDocument sourceDocument, PresentationDocument newDocument, List<ImageData> images, List<MediaData> mediaList)
        {
            XDocument newPresentation = newDocument.PresentationPart.GetXDocument();

            // Copy slide and note slide sizes
            XDocument oldPresentationDoc = sourceDocument.PresentationPart.GetXDocument();

            foreach (var att in oldPresentationDoc.Root.Attributes())
            {
                if (!att.IsNamespaceDeclaration && newPresentation.Root.Attribute(att.Name) == null)
                    newPresentation.Root.Add(oldPresentationDoc.Root.Attribute(att.Name));
            }

            XElement oldElement = oldPresentationDoc.Root.Elements(P.sldSz).FirstOrDefault();
            if (oldElement != null)
                newPresentation.Root.Add(oldElement);

            // Copy Font Parts
            if (oldPresentationDoc.Root.Element(P.embeddedFontLst) != null)
            {
                XElement newFontLst = new XElement(P.embeddedFontLst);
                foreach (var font in oldPresentationDoc.Root.Element(P.embeddedFontLst).Elements(P.embeddedFont))
                {
                    XElement newRegular = null, newBold = null, newItalic = null, newBoldItalic = null;
                    if (font.Element(P.regular) != null)
                        newRegular = CreatedEmbeddedFontPart(sourceDocument, newDocument, font, P.regular);
                    if (font.Element(P.bold) != null)
                        newBold = CreatedEmbeddedFontPart(sourceDocument, newDocument, font, P.bold);
                    if (font.Element(P.italic) != null)
                        newItalic = CreatedEmbeddedFontPart(sourceDocument, newDocument, font, P.italic);
                    if (font.Element(P.boldItalic) != null)
                        newBoldItalic = CreatedEmbeddedFontPart(sourceDocument, newDocument, font, P.boldItalic);
                    XElement newEmbeddedFont = new XElement(P.embeddedFont,
                        font.Elements(P.font),
                        newRegular,
                        newBold,
                        newItalic,
                        newBoldItalic);
                    newFontLst.Add(newEmbeddedFont);
                }
                newPresentation.Root.Add(newFontLst);
            }

            newPresentation.Root.Add(oldPresentationDoc.Root.Element(P.defaultTextStyle));
            newPresentation.Root.Add(oldPresentationDoc.Root.Elements(P.extLst));

            //<p:embeddedFont xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
            //                         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            //  <p:font typeface="Perpetua" panose="02020502060401020303" pitchFamily="18" charset="0" />
            //  <p:regular r:id="rId5" />
            //  <p:bold r:id="rId6" />
            //  <p:italic r:id="rId7" />
            //  <p:boldItalic r:id="rId8" />
            //</p:embeddedFont>

            // Copy Handout Master
            if (sourceDocument.PresentationPart.HandoutMasterPart != null)
            {
                HandoutMasterPart oldMaster = sourceDocument.PresentationPart.HandoutMasterPart;
                HandoutMasterPart newMaster = newDocument.PresentationPart.AddNewPart<HandoutMasterPart>();

                // Copy theme for master
                ThemePart newThemePart = newMaster.AddNewPart<ThemePart>();
                newThemePart.PutXDocument(oldMaster.ThemePart.GetXDocument());
                CopyRelatedPartsForContentParts(newDocument, oldMaster.ThemePart, newThemePart, new[] { newThemePart.GetXDocument().Root }, images, mediaList);

                // Copy master
                newMaster.PutXDocument(oldMaster.GetXDocument());
                AddRelationships(oldMaster, newMaster, new[] { newMaster.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, oldMaster, newMaster, new[] { newMaster.GetXDocument().Root }, images, mediaList);

                newPresentation.Root.Add(
                    new XElement(P.handoutMasterIdLst, new XElement(P.handoutMasterId,
                    new XAttribute(R.id, newDocument.PresentationPart.GetIdOfPart(newMaster)))));
            }

            // Copy Notes Master
            CopyNotesMaster(sourceDocument, newDocument, images, mediaList);

            // Copy Presentation Properties
            if (sourceDocument.PresentationPart.PresentationPropertiesPart != null)
            {
                PresentationPropertiesPart newPart = newDocument.PresentationPart.AddNewPart<PresentationPropertiesPart>();
                XDocument xd1 = sourceDocument.PresentationPart.PresentationPropertiesPart.GetXDocument();
                xd1.Descendants(P.custShow).Remove();
                newPart.PutXDocument(xd1);
            }

            // Copy View Properties
            if (sourceDocument.PresentationPart.ViewPropertiesPart != null)
            {
                ViewPropertiesPart newPart = newDocument.PresentationPart.AddNewPart<ViewPropertiesPart>();
                XDocument xd = sourceDocument.PresentationPart.ViewPropertiesPart.GetXDocument();
                xd.Descendants(P.outlineViewPr).Elements(P.sldLst).Remove();
                newPart.PutXDocument(xd);
            }

            foreach (var legacyDocTextInfo in sourceDocument.PresentationPart.Parts.Where(p => p.OpenXmlPart.RelationshipType == "http://schemas.microsoft.com/office/2006/relationships/legacyDocTextInfo"))
            {
                LegacyDiagramTextInfoPart newPart = newDocument.PresentationPart.AddNewPart<LegacyDiagramTextInfoPart>();
                newPart.FeedData(legacyDocTextInfo.OpenXmlPart.GetStream());
            }

            var listOfRootChildren = newPresentation.Root.Elements().ToList();
            foreach (var rc in listOfRootChildren)
                rc.Remove();
            newPresentation.Root.Add(
                listOfRootChildren.OrderBy(e =>
                {
                    if (_orderPresentation.ContainsKey(e.Name))
                        return _orderPresentation[e.Name];
                    return 999;
                }));
        }

        private static readonly Dictionary<XName, int> _orderPresentation = new()
        {
            { P.sldMasterIdLst, 10 },
            { P.notesMasterIdLst, 20 },
            { P.handoutMasterIdLst, 30 },
            { P.sldIdLst, 40 },
            { P.sldSz, 50 },
            { P.notesSz, 60 },
            { P.embeddedFontLst, 70 },
            { P.custShowLst, 80 },
            { P.photoAlbum, 90 },
            { P.custDataLst, 100 },
            { P.kinsoku, 120 },
            { P.defaultTextStyle, 130 },
            { P.modifyVerifier, 150 },
            { P.extLst, 160 },
        };


        private static XElement CreatedEmbeddedFontPart(PresentationDocument sourceDocument, PresentationDocument newDocument, XElement font, XName fontXName)
        {
            XElement newRegular;
            FontPart oldFontPart = (FontPart)sourceDocument.PresentationPart.GetPartById((string)font.Element(fontXName).Attributes(R.id).FirstOrDefault());
            PartTypeInfo fpt;
            if (oldFontPart.ContentType == "application/x-fontdata")
                fpt = FontPartType.FontData;
            else if (oldFontPart.ContentType == "application/x-font-ttf")
                fpt = FontPartType.FontTtf;
            else
                fpt = FontPartType.FontOdttf;
            var newId = "R" + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 16);
            var newFontPart = newDocument.PresentationPart.AddFontPart(fpt, newId);
            newFontPart.FeedData(oldFontPart.GetStream());
            newRegular = new XElement(fontXName,
                new XAttribute(R.id, newId));
            return newRegular;
        }

        private static SlideMasterPart AppendSlides(PresentationDocument sourceDocument, PresentationDocument newDocument,
            int start, int count, bool keepMaster, List<ImageData> images, SlideMasterPart currentMasterPart, List<MediaData> mediaList)
        {
            XDocument newPresentation = newDocument.PresentationPart.GetXDocument();
            if (newPresentation.Root.Element(P.sldIdLst) == null)
                newPresentation.Root.Add(new XElement(P.sldIdLst));
            uint newID = 256;
            var ids = newPresentation.Root.Descendants(P.sldId).Select(f => (uint)f.Attribute(NoNamespace.id));
            if (ids.Any())
                newID = ids.Max() + 1;
            var slideList = sourceDocument.PresentationPart.GetXDocument().Root.Descendants(P.sldId);
            if (slideList.Count() == 0 && (currentMasterPart == null || keepMaster))
            {
                var slideMasterPart = sourceDocument.PresentationPart.SlideMasterParts.FirstOrDefault();
                if (slideMasterPart != null)
                    currentMasterPart = CopyMasterSlide(sourceDocument, slideMasterPart, newDocument, newPresentation, images, mediaList);
                return currentMasterPart;
            }
            while (count > 0 && start < slideList.Count())
            {
                SlidePart slide = (SlidePart)sourceDocument.PresentationPart.GetPartById(slideList.ElementAt(start).Attribute(R.id).Value);
                if (currentMasterPart == null || keepMaster)
                    currentMasterPart = CopyMasterSlide(sourceDocument, slide.SlideLayoutPart.SlideMasterPart, newDocument, newPresentation, images, mediaList);
                SlidePart newSlide = newDocument.PresentationPart.AddNewPart<SlidePart>();
                newSlide.PutXDocument(slide.GetXDocument());
                AddRelationships(slide, newSlide, new[] { newSlide.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, slide, newSlide, new[] { newSlide.GetXDocument().Root }, images, mediaList);
                CopyTableStyles(sourceDocument, newDocument, slide, newSlide);
                if (slide.NotesSlidePart != null)
                {
                    if (newDocument.PresentationPart.NotesMasterPart == null)
                        CopyNotesMaster(sourceDocument, newDocument, images, mediaList);
                    NotesSlidePart newPart = newSlide.AddNewPart<NotesSlidePart>();
                    newPart.PutXDocument(slide.NotesSlidePart.GetXDocument());
                    newPart.AddPart(newSlide);
                    newPart.AddPart(newDocument.PresentationPart.NotesMasterPart);
                    AddRelationships(slide.NotesSlidePart, newPart, new[] { newPart.GetXDocument().Root });
                    CopyRelatedPartsForContentParts(newDocument, slide.NotesSlidePart, newPart, new[] { newPart.GetXDocument().Root }, images, mediaList);
                }

                string layoutName = slide.SlideLayoutPart.GetXDocument().Root.Element(P.cSld).Attribute(NoNamespace.name).Value;
                foreach (SlideLayoutPart layoutPart in currentMasterPart.SlideLayoutParts)
                    if (layoutPart.GetXDocument().Root.Element(P.cSld).Attribute(NoNamespace.name).Value == layoutName)
                    {
                        newSlide.AddPart(layoutPart);
                        break;
                    }
                if (newSlide.SlideLayoutPart == null)
                    newSlide.AddPart(currentMasterPart.SlideLayoutParts.First());  // Cannot find matching layout part

                if (slide.SlideCommentsPart != null)
                    CopyComments(sourceDocument, newDocument, slide, newSlide);

                newPresentation.Root.Element(P.sldIdLst).Add(new XElement(P.sldId,
                    new XAttribute(NoNamespace.id, newID.ToString()),
                    new XAttribute(R.id, newDocument.PresentationPart.GetIdOfPart(newSlide))));
                newID++;
                start++;
                count--;
            }
            return currentMasterPart;
        }

        private static SlideMasterPart CopyMasterSlide(PresentationDocument sourceDocument, SlideMasterPart sourceMasterPart,
            PresentationDocument newDocument, XDocument newPresentation, List<ImageData> images, List<MediaData> mediaList)
        {
            // Search for existing master slide with same theme name
            var oldTheme = sourceMasterPart.ThemePart.GetXDocument();
            var themeName = oldTheme.Root.Attribute(NoNamespace.name).Value;
            foreach (SlideMasterPart master in newDocument.PresentationPart.GetPartsOfType<SlideMasterPart>())
            {
                var themeDoc = master.ThemePart.GetXDocument();
                if (themeDoc.Root.Attribute(NoNamespace.name).Value == themeName)
                    return master;
            }

            var newMaster = newDocument.PresentationPart.AddNewPart<SlideMasterPart>();
            var sourceMaster = sourceMasterPart.GetXDocument();

            // Add to presentation slide master list, need newID for layout IDs also
            uint newID = 2147483648;
            var ids = newPresentation.Root.Descendants(P.sldMasterId).Select(f => (uint)f.Attribute(NoNamespace.id));
            if (ids.Any())
            {
                newID = ids.Max();
                XElement maxMaster = newPresentation.Root.Descendants(P.sldMasterId).Where(f => (uint)f.Attribute(NoNamespace.id) == newID).FirstOrDefault();
                SlideMasterPart maxMasterPart = (SlideMasterPart)newDocument.PresentationPart.GetPartById(maxMaster.Attribute(R.id).Value);
                newID += (uint)maxMasterPart.GetXDocument().Root.Descendants(P.sldLayoutId).Count() + 1;
            }
            newPresentation.Root.Element(P.sldMasterIdLst).Add(new XElement(P.sldMasterId,
                new XAttribute(NoNamespace.id, newID.ToString()),
                new XAttribute(R.id, newDocument.PresentationPart.GetIdOfPart(newMaster))));
            newID++;

            var newThemePart = newMaster.AddNewPart<ThemePart>();
            if (newDocument.PresentationPart.ThemePart == null)
                newThemePart = newDocument.PresentationPart.AddPart(newThemePart);
            newThemePart.PutXDocument(oldTheme);
            CopyRelatedPartsForContentParts(newDocument, sourceMasterPart.ThemePart, newThemePart, new[] { newThemePart.GetXDocument().Root }, images, mediaList);
            foreach (SlideLayoutPart layoutPart in sourceMasterPart.SlideLayoutParts)
            {
                SlideLayoutPart newLayout = newMaster.AddNewPart<SlideLayoutPart>();
                newLayout.PutXDocument(layoutPart.GetXDocument());
                AddRelationships(layoutPart, newLayout, new[] { newLayout.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, layoutPart, newLayout, new[] { newLayout.GetXDocument().Root }, images, mediaList);
                newLayout.AddPart(newMaster);
                string resID = sourceMasterPart.GetIdOfPart(layoutPart);
                XElement entry = sourceMaster.Root.Descendants(P.sldLayoutId).Where(f => f.Attribute(R.id).Value == resID).FirstOrDefault();
                entry.Attribute(R.id).SetValue(newMaster.GetIdOfPart(newLayout));
                entry.SetAttributeValue(NoNamespace.id, newID.ToString());
                newID++;
            }
            newMaster.PutXDocument(sourceMaster);
            AddRelationships(sourceMasterPart, newMaster, new[] { newMaster.GetXDocument().Root });
            CopyRelatedPartsForContentParts(newDocument, sourceMasterPart, newMaster, new[] { newMaster.GetXDocument().Root }, images, mediaList);

            return newMaster;
        }

        // Copies notes master and notesSz element from presentation
        private static void CopyNotesMaster(PresentationDocument sourceDocument, PresentationDocument newDocument, List<ImageData> images, List<MediaData> mediaList)
        {
            // Copy notesSz element from presentation
            XDocument newPresentation = newDocument.PresentationPart.GetXDocument();
            XDocument oldPresentationDoc = sourceDocument.PresentationPart.GetXDocument();
            XElement oldElement = oldPresentationDoc.Root.Element(P.notesSz);
            newPresentation.Root.Element(P.notesSz).ReplaceWith(oldElement);

            // Copy Notes Master
            if (sourceDocument.PresentationPart.NotesMasterPart != null)
            {
                NotesMasterPart oldMaster = sourceDocument.PresentationPart.NotesMasterPart;
                NotesMasterPart newMaster = newDocument.PresentationPart.AddNewPart<NotesMasterPart>();

                // Copy theme for master
                if (oldMaster.ThemePart != null)
                {
                    ThemePart newThemePart = newMaster.AddNewPart<ThemePart>();
                    newThemePart.PutXDocument(oldMaster.ThemePart.GetXDocument());
                    CopyRelatedPartsForContentParts(newDocument, oldMaster.ThemePart, newThemePart, new[] { newThemePart.GetXDocument().Root }, images, mediaList);
                }

                // Copy master
                newMaster.PutXDocument(oldMaster.GetXDocument());
                AddRelationships(oldMaster, newMaster, new[] { newMaster.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, oldMaster, newMaster, new[] { newMaster.GetXDocument().Root }, images, mediaList);

                newPresentation.Root.Add(
                    new XElement(P.notesMasterIdLst, new XElement(P.notesMasterId,
                    new XAttribute(R.id, newDocument.PresentationPart.GetIdOfPart(newMaster)))));
            }
        }

        private static void CopyComments(PresentationDocument oldDocument, PresentationDocument newDocument, SlidePart oldSlide, SlidePart newSlide)
        {
            newSlide.AddNewPart<SlideCommentsPart>();
            newSlide.SlideCommentsPart.PutXDocument(oldSlide.SlideCommentsPart.GetXDocument());
            XDocument newSlideComments = newSlide.SlideCommentsPart.GetXDocument();
            XDocument oldAuthors = oldDocument.PresentationPart.CommentAuthorsPart.GetXDocument();
            foreach (XElement comment in newSlideComments.Root.Elements(P.cm))
            {
                XElement newAuthor = FindCommentsAuthor(newDocument, comment, oldAuthors);
                // Update last index value for new comment
                comment.Attribute(NoNamespace.authorId).SetValue(newAuthor.Attribute(NoNamespace.id).Value);
                uint lastIndex = Convert.ToUInt32(newAuthor.Attribute(NoNamespace.lastIdx).Value);
                comment.Attribute(NoNamespace.idx).SetValue(lastIndex.ToString());
                newAuthor.Attribute(NoNamespace.lastIdx).SetValue(Convert.ToString(lastIndex + 1));
            }
        }

        private static XElement FindCommentsAuthor(PresentationDocument newDocument, XElement comment, XDocument oldAuthors)
        {
            XElement oldAuthor = oldAuthors.Root.Elements(P.cmAuthor).Where(
                f => f.Attribute(NoNamespace.id).Value == comment.Attribute(NoNamespace.authorId).Value).FirstOrDefault();
            XElement newAuthor = null;
            if (newDocument.PresentationPart.CommentAuthorsPart == null)
            {
                newDocument.PresentationPart.AddNewPart<CommentAuthorsPart>();
                newDocument.PresentationPart.CommentAuthorsPart.PutXDocument(new XDocument(new XElement(P.cmAuthorLst,
                    new XAttribute(XNamespace.Xmlns + "a", A.a),
                    new XAttribute(XNamespace.Xmlns + "r", R.r),
                    new XAttribute(XNamespace.Xmlns + "p", P.p))));
            }
            XDocument authors = newDocument.PresentationPart.CommentAuthorsPart.GetXDocument();
            newAuthor = authors.Root.Elements(P.cmAuthor).Where(
                f => f.Attribute(NoNamespace.initials).Value == oldAuthor.Attribute(NoNamespace.initials).Value).FirstOrDefault();
            if (newAuthor == null)
            {
                uint newID = 0;
                var ids = authors.Root.Descendants(P.cmAuthor).Select(f => (uint)f.Attribute(NoNamespace.id));
                if (ids.Any())
                    newID = ids.Max() + 1;

                newAuthor = new XElement(P.cmAuthor, new XAttribute(NoNamespace.id, newID.ToString()),
                    new XAttribute(NoNamespace.name, oldAuthor.Attribute(NoNamespace.name).Value),
                    new XAttribute(NoNamespace.initials, oldAuthor.Attribute(NoNamespace.initials).Value),
                    new XAttribute(NoNamespace.lastIdx, "1"), new XAttribute(NoNamespace.clrIdx, newID.ToString()));
                authors.Root.Add(newAuthor);
            }

            return newAuthor;
        }

        private static void CopyTableStyles(PresentationDocument oldDocument, PresentationDocument newDocument, OpenXmlPart oldContentPart, OpenXmlPart newContentPart)
        {
            foreach (XElement table in newContentPart.GetXDocument().Descendants(A.tableStyleId))
            {
                string styleId = table.Value;
                if (string.IsNullOrEmpty(styleId))
                    continue;

                // Find old style
                if (oldDocument.PresentationPart.TableStylesPart == null)
                    continue;
                XDocument oldTableStyles = oldDocument.PresentationPart.TableStylesPart.GetXDocument();
                XElement oldStyle = oldTableStyles.Root.Elements(A.tblStyle).Where(f => f.Attribute(NoNamespace.styleId).Value == styleId).FirstOrDefault();
                if (oldStyle == null)
                    continue;

                // Create new TableStylesPart, if needed
                XDocument tableStyles = null;
                if (newDocument.PresentationPart.TableStylesPart == null)
                {
                    TableStylesPart newStylesPart = newDocument.PresentationPart.AddNewPart<TableStylesPart>();
                    tableStyles = new XDocument(new XElement(A.tblStyleLst,
                        new XAttribute(XNamespace.Xmlns + "a", A.a),
                        new XAttribute(NoNamespace.def, styleId)));
                    newStylesPart.PutXDocument(tableStyles);
                }
                else
                    tableStyles = newDocument.PresentationPart.TableStylesPart.GetXDocument();

                // Search new TableStylesPart to see if it contains the ID
                if (tableStyles.Root.Elements(A.tblStyle).Where(f => f.Attribute(NoNamespace.styleId).Value == styleId).FirstOrDefault() != null)
                    continue;

                // Copy style to new part
                tableStyles.Root.Add(oldStyle);
            }

        }

        private static void CopyRelatedPartsForContentParts(PresentationDocument newDocument, OpenXmlPart oldContentPart, OpenXmlPart newContentPart,
            IEnumerable<XElement> newContent, List<ImageData> images, List<MediaData> mediaList)
        {
            var relevantElements = newContent.DescendantsAndSelf()
                .Where(d => d.Name == VML.imagedata || d.Name == VML.fill || d.Name == VML.stroke || d.Name == A.blip || d.Name == SVG.svgBlip)
                .ToList();
            foreach (XElement imageReference in relevantElements)
            {
                CopyRelatedImage(oldContentPart, newContentPart, imageReference, R.embed, images);
                CopyRelatedImage(oldContentPart, newContentPart, imageReference, R.pict, images);
                CopyRelatedImage(oldContentPart, newContentPart, imageReference, R.id, images);
                CopyRelatedImage(oldContentPart, newContentPart, imageReference, O.relid, images);
            }

            relevantElements = newContent.DescendantsAndSelf()
                .Where(d => d.Name == A.videoFile || d.Name == A.quickTimeFile)
                .ToList();
            foreach (XElement imageReference in relevantElements)
            {
                CopyRelatedMedia(oldContentPart, newContentPart, imageReference, R.link, mediaList, "video");
            }

            relevantElements = newContent.DescendantsAndSelf()
                .Where(d => d.Name == P14.media || d.Name == PAV.srcMedia)
                .ToList();
            foreach (XElement imageReference in relevantElements)
            {
                CopyRelatedMedia(oldContentPart, newContentPart, imageReference, R.embed, mediaList, "media");
                CopyRelatedMediaExternalRelationship(oldContentPart, newContentPart, imageReference, R.link, "media");
            }

            foreach (XElement extendedReference in newContent.DescendantsAndSelf(A14.imgLayer))
            {
                CopyExtendedPart(oldContentPart, newContentPart, extendedReference, R.embed);
            }

            foreach (XElement contentPartReference in newContent.DescendantsAndSelf(P.contentPart))
            {
                CopyInkPart(oldContentPart, newContentPart, contentPartReference, R.id);
            }

            foreach (XElement contentPartReference in newContent.DescendantsAndSelf(P.control))
            {
                CopyActiveXPart(oldContentPart, newContentPart, contentPartReference, R.id);
            }

            foreach (XElement contentPartReference in newContent.DescendantsAndSelf(Plegacy.textdata))
            {
                CopyLegacyDiagramText(oldContentPart, newContentPart, contentPartReference, "id");
            }

            foreach (XElement diagramReference in newContent.DescendantsAndSelf().Where(d => d.Name == DGM.relIds || d.Name == A.relIds))
            {
                // dm attribute
                string relId = diagramReference.Attribute(R.dm).Value;
                var tempPartIdPair = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair != default)
                    continue;

                ExternalRelationship tempEr = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
                if (tempEr != null)
                    continue;

                OpenXmlPart oldPart = oldContentPart.GetPartById(relId);
                OpenXmlPart newPart = newContentPart.AddNewPart<DiagramDataPart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.dm).Value = newContentPart.GetIdOfPart(newPart);
                AddRelationships(oldPart, newPart, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, oldPart, newPart, new[] { newPart.GetXDocument().Root }, images, mediaList);

                // lo attribute
                relId = diagramReference.Attribute(R.lo).Value;
                var tempPartIdPair2 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair2 != default)
                    continue;

                var tempEr2 = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
                if (tempEr2 != null)
                    continue;

                oldPart = oldContentPart.GetPartById(relId);
                newPart = newContentPart.AddNewPart<DiagramLayoutDefinitionPart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.lo).Value = newContentPart.GetIdOfPart(newPart);
                AddRelationships(oldPart, newPart, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, oldPart, newPart, new[] { newPart.GetXDocument().Root }, images, mediaList);

                // qs attribute
                relId = diagramReference.Attribute(R.qs).Value;
                var tempPartIdPair3 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair3 != default)
                    continue;

                ExternalRelationship tempEr3 = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
                if (tempEr3 != null)
                    continue;

                oldPart = oldContentPart.GetPartById(relId);
                newPart = newContentPart.AddNewPart<DiagramStylePart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.qs).Value = newContentPart.GetIdOfPart(newPart);
                AddRelationships(oldPart, newPart, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, oldPart, newPart, new[] { newPart.GetXDocument().Root }, images, mediaList);

                // cs attribute
                relId = diagramReference.Attribute(R.cs).Value;
                var tempPartIdPair4 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair4 != default)
                    continue;

                ExternalRelationship tempEr4 = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
                if (tempEr4 != null)
                    continue;

                oldPart = oldContentPart.GetPartById(relId);
                newPart = newContentPart.AddNewPart<DiagramColorsPart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.cs).Value = newContentPart.GetIdOfPart(newPart);
                AddRelationships(oldPart, newPart, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, oldPart, newPart, new[] { newPart.GetXDocument().Root }, images, mediaList);
            }

            foreach (var oleReference in newContent.DescendantsAndSelf().Where(d => d.Name == P.oleObj || d.Name == P.externalData))
            {
                string relId = oleReference.Attribute(R.id).Value;

                // First look to see if this relId has already been added to the new document.
                // This is necessary for those parts that get processed with both old and new ids, such as the comments
                // part.  This is not necessary for parts such as the main document part, but this code won't malfunction
                // in that case.
                var tempPartIdPair5 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair5 != default)
                    continue;

                ExternalRelationship tempEr5 = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
                if (tempEr5 != null)
                    continue;

                var oldPartIdPair = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (oldPartIdPair != default)
                {
                    var oldPart = oldPartIdPair.OpenXmlPart;
                    OpenXmlPart newPart = null;
                    newPart = oldPart switch
                    {
                        EmbeddedObjectPart => newContentPart switch
                        {
                            DialogsheetPart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                            HandoutMasterPart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                            NotesMasterPart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                            NotesSlidePart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                            SlideLayoutPart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                            SlideMasterPart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                            SlidePart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                            _ => newPart
                        },
                        EmbeddedPackagePart => newContentPart switch
                        {
                            ChartPart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                            HandoutMasterPart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                            NotesMasterPart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                            NotesSlidePart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                            SlideLayoutPart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                            SlideMasterPart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                            SlidePart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                            _ => newPart
                        },
                        _ => newPart
                    };

                    using (var oldObject = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                    {
                        newPart.FeedData(oldObject);
                    }
                    oleReference.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                }
                else
                {
                    var er = oldContentPart.GetExternalRelationship(relId);
                    var newEr = newContentPart.AddExternalRelationship(er.RelationshipType, er.Uri);
                    oleReference.Attribute(R.id).Value = newEr.Id;
                }
            }

            foreach (XElement chartReference in newContent.DescendantsAndSelf(C.chart))
            {
                var relId = (string)chartReference.Attribute(R.id);
                if (string.IsNullOrEmpty(relId))
                    continue;

                var tempPartIdPair6 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair6 != default)
                    continue;

                var tempEr6 = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
                if (tempEr6 != null)
                    continue;

                var oldPartIdPair2 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (oldPartIdPair2 != default)
                {
                    if (oldPartIdPair2.OpenXmlPart is ChartPart oldPart)
                    {
                        XDocument oldChart = oldPart.GetXDocument();
                        ChartPart newPart = newContentPart.AddNewPart<ChartPart>();
                        XDocument newChart = newPart.GetXDocument();
                        newChart.Add(oldChart.Root);
                        chartReference.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                        CopyChartObjects(oldPart, newPart);
                        CopyRelatedPartsForContentParts(newDocument, oldPart, newPart, new[] { newChart.Root }, images, mediaList);
                    }
                }
            }

            foreach (XElement userShape in newContent.DescendantsAndSelf(C.userShapes))
            {
                var relId = (string)userShape.Attribute(R.id);
                if (string.IsNullOrEmpty(relId))
                    continue;

                var tempPartIdPair7 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair7 != default)
                    continue;

                var tempEr7 = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
                if (tempEr7 != null)
                    continue;

                var oldPartIdPair3 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (oldPartIdPair3 != default)
                {
                    if (oldPartIdPair3.OpenXmlPart is ChartDrawingPart oldPart)
                    {
                        var oldXDoc = oldPart.GetXDocument();
                        var newPart = newContentPart.AddNewPart<ChartDrawingPart>();
                        var newXDoc = newPart.GetXDocument();
                        newXDoc.Add(oldXDoc.Root);
                        userShape.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                        AddRelationships(oldPart, newPart, newContent);
                        CopyRelatedPartsForContentParts(newDocument, oldPart, newPart, new[] { newXDoc.Root }, images, mediaList);
                    }
                }
            }

            foreach (XElement tags in newContent.DescendantsAndSelf(P.tags))
            {
                var relId = (string)tags.Attribute(R.id);
                if (string.IsNullOrEmpty(relId))
                    continue;

                var tempPartIdPair8 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair8 != default)
                    continue;

                var tempEr8 = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
                if (tempEr8 != null)
                    continue;

                var oldPartIdPair4 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (oldPartIdPair4 != default)
                {
                    if (oldPartIdPair4.OpenXmlPart is UserDefinedTagsPart oldPart)
                    {
                        var oldXDoc = oldPart.GetXDocument();
                        var newPart = newContentPart.AddNewPart<UserDefinedTagsPart>();
                        var newXDoc = newPart.GetXDocument();
                        newXDoc.Add(oldXDoc.Root);
                        tags.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                    }
                }
            }

            foreach (XElement custData in newContent.DescendantsAndSelf(P.custData))
            {
                var relId = (string)custData.Attribute(R.id);
                if (string.IsNullOrEmpty(relId))
                    continue;

                var tempPartIdPair9 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair9 != default)
                    continue;

                var oldPartIdPair9 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (oldPartIdPair9 != default)
                {
                    var newPart = newDocument.PresentationPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                    newPart.FeedData(oldPartIdPair9.OpenXmlPart.GetStream());
                    foreach (var itemProps in oldPartIdPair9.OpenXmlPart.Parts.Where(p => p.OpenXmlPart.ContentType == "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"))
                    {
                        var newId2 = "R" + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 16);
                        var cxpp = newPart.AddNewPart<CustomXmlPropertiesPart>("application/vnd.openxmlformats-officedocument.customXmlProperties+xml", newId2);
                        cxpp.FeedData(itemProps.OpenXmlPart.GetStream());
                    }
                    var newId = "R" + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 16);
                    newContentPart.CreateRelationshipToPart(newPart, newId);
                    custData.Attribute(R.id).Value = newId;
                }
            }

            foreach (XElement soundReference in newContent.DescendantsAndSelf().Where(d => d.Name == A.audioFile))
                CopyRelatedSound(newDocument, oldContentPart, newContentPart, soundReference, R.link);

            if ((oldContentPart is ChartsheetPart && newContentPart is ChartsheetPart) ||
                (oldContentPart is DialogsheetPart && newContentPart is DialogsheetPart) ||
                (oldContentPart is HandoutMasterPart && newContentPart is HandoutMasterPart) ||
                (oldContentPart is InternationalMacroSheetPart && newContentPart is InternationalMacroSheetPart) ||
                (oldContentPart is MacroSheetPart && newContentPart is MacroSheetPart) ||
                (oldContentPart is NotesMasterPart && newContentPart is NotesMasterPart) ||
                (oldContentPart is NotesSlidePart && newContentPart is NotesSlidePart) ||
                (oldContentPart is SlideLayoutPart && newContentPart is SlideLayoutPart) ||
                (oldContentPart is SlideMasterPart && newContentPart is SlideMasterPart) ||
                (oldContentPart is SlidePart && newContentPart is SlidePart) ||
                (oldContentPart is WorksheetPart && newContentPart is WorksheetPart))
            {
                foreach (XElement soundReference in newContent.DescendantsAndSelf().Where(d => d.Name == P.snd || d.Name == P.sndTgt || d.Name == A.wavAudioFile || d.Name == A.snd || d.Name == PAV.srcMedia))
                    CopyRelatedSound(newDocument, oldContentPart, newContentPart, soundReference, R.embed);

                var vmlDrawingParts = oldContentPart switch
                {
                    ChartsheetPart part => part.VmlDrawingParts,
                    DialogsheetPart part => part.VmlDrawingParts,
                    HandoutMasterPart part => part.VmlDrawingParts,
                    InternationalMacroSheetPart part => part.VmlDrawingParts,
                    MacroSheetPart part => part.VmlDrawingParts,
                    NotesMasterPart part => part.VmlDrawingParts,
                    NotesSlidePart part => part.VmlDrawingParts,
                    SlideLayoutPart part => part.VmlDrawingParts,
                    SlideMasterPart part => part.VmlDrawingParts,
                    SlidePart part => part.VmlDrawingParts,
                    WorksheetPart part => part.VmlDrawingParts,
                    _ => null
                };

                if (vmlDrawingParts != null)
                {
                    // Transitional: Copy VML Drawing parts, implicit relationship
                    foreach (VmlDrawingPart vmlPart in vmlDrawingParts)
                    {
                        var newVmlPart = newContentPart switch
                        {
                            ChartsheetPart part => part.AddNewPart<VmlDrawingPart>(),
                            DialogsheetPart part => part.AddNewPart<VmlDrawingPart>(),
                            HandoutMasterPart part => part.AddNewPart<VmlDrawingPart>(),
                            InternationalMacroSheetPart part => part.AddNewPart<VmlDrawingPart>(),
                            MacroSheetPart part => part.AddNewPart<VmlDrawingPart>(),
                            NotesMasterPart part => part.AddNewPart<VmlDrawingPart>(),
                            NotesSlidePart part => part.AddNewPart<VmlDrawingPart>(),
                            SlideLayoutPart part => part.AddNewPart<VmlDrawingPart>(),
                            SlideMasterPart part => part.AddNewPart<VmlDrawingPart>(),
                            SlidePart part => part.AddNewPart<VmlDrawingPart>(),
                            WorksheetPart part => part.AddNewPart<VmlDrawingPart>(),
                            _ => null
                        };

                        var xd = vmlPart.GetXDocument();
                        foreach (var item in xd.Descendants(O.ink))
                        {
                            if (item.Attribute("i") != null)
                            {
                                var i = item.Attribute("i").Value;
                                i = i.Replace(" ", "\r\n");
                                item.Attribute("i").Value = i;
                            }
                        }
                        newVmlPart.PutXDocument(xd);
                        AddRelationships(vmlPart, newVmlPart, new[] { newVmlPart.GetXDocument().Root });
                        CopyRelatedPartsForContentParts(newDocument, vmlPart, newVmlPart, new[] { newVmlPart.GetXDocument().Root }, images, mediaList);
                    }
                }
            }
        }

        private static void CopyChartObjects(ChartPart oldChart, ChartPart newChart)
        {
            foreach (var dataReference in newChart.GetXDocument().Descendants(C.externalData))
            {
                string relId = dataReference.Attribute(R.id).Value;

                var oldPartIdPair = oldChart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (oldPartIdPair != default)
                {
                    switch (oldPartIdPair.OpenXmlPart)
                    {
                        case EmbeddedPackagePart oldPart:
                            {
                                var newPart = newChart.AddEmbeddedPackagePart(oldPart.ContentType);
                                using (var oldObject = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                                {
                                    newPart.FeedData(oldObject);
                                }
                                dataReference.Attribute(R.id).Value = newChart.GetIdOfPart(newPart);
                                continue;
                            }

                        case EmbeddedObjectPart oldEmbeddedObjectPart:
                            {
                                var newPart = newChart.AddEmbeddedPackagePart(oldEmbeddedObjectPart.ContentType);
                                using (var oldObject = oldEmbeddedObjectPart.GetStream(FileMode.Open, FileAccess.Read))
                                {
                                    newPart.FeedData(oldObject);
                                }

                                var rId = newChart.GetIdOfPart(newPart);
                                dataReference.Attribute(R.id).Value = rId;

                                // following is a hack to fix the package because the Open XML SDK does not let us create
                                // a relationship from a chart with the oleObject relationship type.

                                var pkg = newChart.OpenXmlPackage.GetPackage();
                                var fromPart = pkg.GetParts().FirstOrDefault(p => p.Uri == newChart.Uri);
                                if (fromPart is not null)
                                {
                                    var rel = fromPart.Relationships.FirstOrDefault(p => p.Id == rId);
                                    var targetUri = rel?.TargetUri;

                                    fromPart.Relationships.Remove(rId);
                                    fromPart.Relationships.Create(targetUri, System.IO.Packaging.TargetMode.Internal,
                                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
                                        rId);
                                }
                                continue;
                            }
                    }
                }
                else
                {
                    ExternalRelationship oldRelationship = oldChart.GetExternalRelationship(relId);
                    Guid g = Guid.NewGuid();
                    string newRid = "R" + g.ToString().Replace("-", "");
                    var oldRel = oldChart.ExternalRelationships.FirstOrDefault(h => h.Id == relId)
                        ?? throw new PresentationBuilderInternalException("Internal Error 0007");

                    newChart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    dataReference.Attribute(R.id).Value = newRid;
                }
            }
        }

        private static Dictionary<XName, XName[]> _relationshipMarkup = null;

        private static void UpdateContent(IEnumerable<XElement> newContent, XName elementToModify, string oldRid, string newRid)
        {
            foreach (var attributeName in _relationshipMarkup[elementToModify])
            {
                var elementsToUpdate = newContent
                    .Descendants(elementToModify)
                    .Where(e => (string)e.Attribute(attributeName) == oldRid);
                foreach (var element in elementsToUpdate)
                    element.Attribute(attributeName).Value = newRid;
            }
        }

        private static void RemoveContent(IEnumerable<XElement> newContent, XName elementToModify, string oldRid)
        {
            foreach (var attributeName in _relationshipMarkup[elementToModify])
            {
                newContent
                    .Descendants(elementToModify)
                    .Where(e => (string)e.Attribute(attributeName) == oldRid).Remove();
            }
        }

        private static void AddRelationships(OpenXmlPart oldPart, OpenXmlPart newPart, IEnumerable<XElement> newContent)
        {
            var relevantElements = newContent.DescendantsAndSelf()
                .Where(d => _relationshipMarkup.ContainsKey(d.Name) &&
                    d.Attributes().Any(a => _relationshipMarkup[d.Name].Contains(a.Name)))
                .ToList();
            foreach (var e in relevantElements)
            {
                if (e.Name == A.hlinkClick || e.Name == A.hlinkHover || e.Name == A.hlinkMouseOver)
                {
                    string relId = (string)e.Attribute(R.id);
                    if (string.IsNullOrEmpty(relId))
                    {
                        // handle the following:
                        //<a:hlinkClick r:id=""
                        //              action="ppaction://customshow?id=0" />
                        var action = (string)e.Attribute("action");
                        if (action != null)
                        {
                            if (action.Contains("customshow"))
                                e.Attribute("action").Remove();
                        }
                        continue;
                    }
                    var tempHyperlink = newPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempHyperlink != null)
                        continue;
                    Guid g = Guid.NewGuid();
                    string newRid = "R" + g.ToString().Replace("-", "");
                    var oldHyperlink = oldPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldHyperlink == null)
                    {
                        //TODO Issue with reference to another part: var temp = oldPart.GetPartById(relId);
                        RemoveContent(newContent, e.Name, relId);
                        continue;
                    }
                    newPart.AddHyperlinkRelationship(oldHyperlink.Uri, oldHyperlink.IsExternal, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
                if (e.Name == VML.imagedata)
                {
                    string relId = (string)e.Attribute(R.href);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempExternalRelationship = newPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempExternalRelationship != null)
                        continue;
                    Guid g = Guid.NewGuid();
                    string newRid = "R" + g.ToString().Replace("-", "");
                    var oldRel = oldPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel == null)
                        throw new PresentationBuilderInternalException("Internal Error 0006");
                    newPart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
                if (e.Name == A.blip || e.Name == A14.imgLayer || e.Name == A.audioFile || e.Name == A.videoFile || e.Name == A.quickTimeFile)
                {
                    string relId = (string)e.Attribute(R.link);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempExternalRelationship = newPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempExternalRelationship != null)
                        continue;
                    Guid g = Guid.NewGuid();
                    string newRid = "R" + g.ToString().Replace("-", "");
                    var oldRel = oldPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel == null)
                        continue;
                    newPart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
            }
        }

        private static void CopyRelatedImage(OpenXmlPart oldContentPart,
                                             OpenXmlPart newContentPart,
                                             XElement imageReference,
                                             XName attributeName,
                                             List<ImageData> images)
        {
            var relId = (string)imageReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;

            // First look to see if this relId has already been added to the new document.
            // This is necessary for those parts that get processed with both old and new ids, such as the comments
            // part.  This is not necessary for parts such as the main document part, but this code won't malfunction
            // in that case.
            var partIdPair = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
            if (partIdPair != default)
                return;

            var extRel = newContentPart.ExternalRelationships.FirstOrDefault(r => r.Id == relId);
            if (extRel != null)
                return;

            var oldPartIdPair = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
            if (oldPartIdPair != default)
            {
                ImagePart oldPart = oldPartIdPair.OpenXmlPart as ImagePart;
                ImageData temp = ManageImageCopy(oldPart, images);
                if (temp.ImagePart == null)
                {
                    ImagePart newPart = null;
                    newPart = newContentPart switch
                    {
                        ChartDrawingPart part => part.AddImagePart(oldPart.ContentType),
                        ChartPart part => part.AddImagePart(oldPart.ContentType),
                        ChartsheetPart part => part.AddImagePart(oldPart.ContentType),
                        DiagramDataPart part => part.AddImagePart(oldPart.ContentType),
                        DiagramLayoutDefinitionPart part => part.AddImagePart(oldPart.ContentType),
                        DiagramPersistLayoutPart part => part.AddImagePart(oldPart.ContentType),
                        DrawingsPart part => part.AddImagePart(oldPart.ContentType),
                        HandoutMasterPart part => part.AddImagePart(oldPart.ContentType),
                        NotesMasterPart part => part.AddImagePart(oldPart.ContentType),
                        NotesSlidePart part => part.AddImagePart(oldPart.ContentType),
                        RibbonAndBackstageCustomizationsPart part => part.AddImagePart(oldPart.ContentType),
                        RibbonExtensibilityPart part => part.AddImagePart(oldPart.ContentType),
                        SlideLayoutPart part => part.AddImagePart(oldPart.ContentType),
                        SlideMasterPart part => part.AddImagePart(oldPart.ContentType),
                        SlidePart part => part.AddImagePart(oldPart.ContentType),
                        ThemeOverridePart part => part.AddImagePart(oldPart.ContentType),
                        ThemePart part => part.AddImagePart(oldPart.ContentType),
                        VmlDrawingPart part => part.AddImagePart(oldPart.ContentType),
                        WorksheetPart part => part.AddImagePart(oldPart.ContentType),
                        _ => newPart,
                    };
                    temp.ImagePart = newPart;
                    var id = newContentPart.GetIdOfPart(newPart);
                    temp.AddContentPartRelTypeResourceIdTupple(newContentPart, newPart.RelationshipType, id);

                    temp.WriteImage(newPart);
                    imageReference.Attribute(attributeName).Value = id;
                }
                else
                {
                    var refRel = newContentPart.DataPartReferenceRelationships.FirstOrDefault(rr =>
                        {
                            var rel = temp.ContentPartRelTypeIdList.FirstOrDefault(cpr =>
                            {
                                var found = cpr.ContentPart == newContentPart && cpr.RelationshipId == rr.Id;
                                return found;
                            });
                            if (rel != null)
                                return true;
                            return false;
                        });
                    if (refRel != null)
                    {
                        imageReference.Attribute(attributeName).Value = temp.ContentPartRelTypeIdList.First(cpr =>
                        {
                            var found = cpr.ContentPart == newContentPart && cpr.RelationshipId == refRel.Id;
                            return found;
                        }).RelationshipId;
                        return;
                    }

                    var cpr2 = temp.ContentPartRelTypeIdList.FirstOrDefault(c => c.ContentPart == newContentPart);
                    if (cpr2 != null)
                    {
                        imageReference.Attribute(attributeName).Value = cpr2.RelationshipId;
                    }
                    else
                    {
                        ImagePart imagePart = (ImagePart)temp.ImagePart;
                        var existingImagePart = newContentPart.AddPart<ImagePart>(imagePart);
                        var newId = newContentPart.GetIdOfPart(existingImagePart);
                        temp.AddContentPartRelTypeResourceIdTupple(newContentPart, imagePart.RelationshipType, newId);
                        imageReference.Attribute(attributeName).Value = newId;
                    }

                }
            }
            else
            {
                ExternalRelationship er = oldContentPart.ExternalRelationships.FirstOrDefault(r => r.Id == relId);
                if (er != null)
                {
                    ExternalRelationship newEr = newContentPart.AddExternalRelationship(er.RelationshipType, er.Uri);
                    imageReference.Attribute(R.id).Value = newEr.Id;
                }
                else
                {
                    var newPart = newContentPart.OpenXmlPackage.GetPackage().GetParts().FirstOrDefault(p => p.Uri == newContentPart.Uri);
                    if (newPart is not null && !newPart.Relationships.Contains(relId))
                    {
                        newPart.Relationships.Create(new Uri("NULL", UriKind.RelativeOrAbsolute),
                            System.IO.Packaging.TargetMode.Internal,
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", relId);
                    }
                }
            }
        }

        private static void CopyRelatedMedia(OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement imageReference, XName attributeName,
            List<MediaData> mediaList, string mediaRelationshipType)
        {
            string relId = (string)imageReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;

            // First look to see if this relId has already been added to the new document.
            var existingDataPartRefRel2 = newContentPart.DataPartReferenceRelationships.FirstOrDefault(dpr => dpr.Id == relId);
            if (existingDataPartRefRel2 != null)
                return;

            var oldRel = oldContentPart.DataPartReferenceRelationships.FirstOrDefault(dpr => dpr.Id == relId);
            if (oldRel == null)
                return;

            DataPart oldPart = oldRel.DataPart;
            MediaData temp = ManageMediaCopy(oldPart, mediaList);
            if (temp.DataPart == null)
            {
                var ct = oldPart.ContentType;
                var ext = Path.GetExtension(oldPart.Uri.OriginalString);
                MediaDataPart newPart = newContentPart.OpenXmlPackage.CreateMediaDataPart(ct, ext);
                newPart.FeedData(oldPart.GetStream());
                string id = null;
                string relationshipType = null;

                if (mediaRelationshipType == "media")
                {
                    MediaReferenceRelationship mrr = null;

                    if (newContentPart is SlidePart)
                        mrr = ((SlidePart)newContentPart).AddMediaReferenceRelationship(newPart);
                    else if (newContentPart is SlideLayoutPart)
                        mrr = ((SlideLayoutPart)newContentPart).AddMediaReferenceRelationship(newPart);
                    else if (newContentPart is SlideMasterPart)
                        mrr = ((SlideMasterPart)newContentPart).AddMediaReferenceRelationship(newPart);

                    id = mrr.Id;
                    relationshipType = "http://schemas.microsoft.com/office/2007/relationships/media";
                }
                else if (mediaRelationshipType == "video")
                {
                    VideoReferenceRelationship vrr = null;

                    if (newContentPart is SlidePart)
                        vrr = ((SlidePart)newContentPart).AddVideoReferenceRelationship(newPart);
                    else if (newContentPart is HandoutMasterPart)
                        vrr = ((HandoutMasterPart)newContentPart).AddVideoReferenceRelationship(newPart);
                    else if (newContentPart is NotesMasterPart)
                        vrr = ((NotesMasterPart)newContentPart).AddVideoReferenceRelationship(newPart);
                    else if (newContentPart is NotesSlidePart)
                        vrr = ((NotesSlidePart)newContentPart).AddVideoReferenceRelationship(newPart);
                    else if (newContentPart is SlideLayoutPart)
                        vrr = ((SlideLayoutPart)newContentPart).AddVideoReferenceRelationship(newPart);
                    else if (newContentPart is SlideMasterPart)
                        vrr = ((SlideMasterPart)newContentPart).AddVideoReferenceRelationship(newPart);

                    id = vrr.Id;
                    relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video";
                }
                temp.DataPart = newPart;
                temp.AddContentPartRelTypeResourceIdTupple(newContentPart, relationshipType, id);
                imageReference.Attribute(attributeName).Value = id;
            }
            else
            {
                string desiredRelType = null;
                if (mediaRelationshipType == "media")
                    desiredRelType = "http://schemas.microsoft.com/office/2007/relationships/media";
                if (mediaRelationshipType == "video")
                    desiredRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video";
                var existingRel = temp.ContentPartRelTypeIdList.FirstOrDefault(cp => cp.ContentPart == newContentPart && cp.RelationshipType == desiredRelType);
                if (existingRel != null)
                {
                    imageReference.Attribute(attributeName).Value = existingRel.RelationshipId;
                }
                else
                {
                    MediaDataPart newPart = (MediaDataPart)temp.DataPart;
                    string id = null;
                    string relationshipType = null;
                    if (mediaRelationshipType == "media")
                    {
                        MediaReferenceRelationship mrr = null;

                        if (newContentPart is SlidePart)
                            mrr = ((SlidePart)newContentPart).AddMediaReferenceRelationship(newPart);
                        if (newContentPart is SlideLayoutPart)
                            mrr = ((SlideLayoutPart)newContentPart).AddMediaReferenceRelationship(newPart);
                        if (newContentPart is SlideMasterPart)
                            mrr = ((SlideMasterPart)newContentPart).AddMediaReferenceRelationship(newPart);

                        id = mrr.Id;
                        relationshipType = mrr.RelationshipType;
                    }
                    else if (mediaRelationshipType == "video")
                    {
                        VideoReferenceRelationship vrr = null;

                        if (newContentPart is SlidePart)
                            vrr = ((SlidePart)newContentPart).AddVideoReferenceRelationship(newPart);
                        if (newContentPart is HandoutMasterPart)
                            vrr = ((HandoutMasterPart)newContentPart).AddVideoReferenceRelationship(newPart);
                        if (newContentPart is NotesMasterPart)
                            vrr = ((NotesMasterPart)newContentPart).AddVideoReferenceRelationship(newPart);
                        if (newContentPart is NotesSlidePart)
                            vrr = ((NotesSlidePart)newContentPart).AddVideoReferenceRelationship(newPart);
                        if (newContentPart is SlideLayoutPart)
                            vrr = ((SlideLayoutPart)newContentPart).AddVideoReferenceRelationship(newPart);
                        if (newContentPart is SlideMasterPart)
                            vrr = ((SlideMasterPart)newContentPart).AddVideoReferenceRelationship(newPart);

                        id = vrr.Id;
                        relationshipType = vrr.RelationshipType;
                    }
                    temp.AddContentPartRelTypeResourceIdTupple(newContentPart, relationshipType, id);
                    imageReference.Attribute(attributeName).Value = id;
                }
            }
        }

        private static void CopyRelatedMediaExternalRelationship(OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement imageReference, XName attributeName,
            string mediaRelationshipType)
        {
            string relId = (string)imageReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;

            var existingExternalReference = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
            if (existingExternalReference != null)
                return;

            var oldRel = oldContentPart.ExternalRelationships.FirstOrDefault(dpr => dpr.Id == relId);
            if (oldRel == null)
                return;

            var newId = "R" + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 16);
            newContentPart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newId);

            imageReference.Attribute(attributeName).Value = newId;
        }


        private static void CopyInkPart(OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement contentPartReference, XName attributeName)
        {
            string relId = (string)contentPartReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;

            var tempPartIdPair = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
            if (tempPartIdPair != default)
                return;

            var tempEr = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
            if (tempEr != null)
                return;

            var oldPart = oldContentPart.GetPartById(relId);

            var newId = "R" + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 16);
            CustomXmlPart newPart = newContentPart.AddNewPart<CustomXmlPart>("application/inkml+xml", newId);

            newPart.FeedData(oldPart.GetStream());
            contentPartReference.Attribute(attributeName).Value = newId;
        }

        private static void CopyActiveXPart(OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement activeXPartReference, XName attributeName)
        {
            string relId = (string)activeXPartReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;

            var tempPartIdPair = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
            if (tempPartIdPair != default)
                return;

            var oldPart = oldContentPart.GetPartById(relId);

            var newId = "R" + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 16);
            EmbeddedControlPersistencePart newPart = newContentPart.AddNewPart<EmbeddedControlPersistencePart>("application/vnd.ms-office.activeX+xml", newId);

            newPart.FeedData(oldPart.GetStream());
            activeXPartReference.Attribute(attributeName).Value = newId;

            if (newPart.ContentType == "application/vnd.ms-office.activeX+xml")
            {
                XDocument axc = newPart.GetXDocument();
                if (axc.Root.Attribute(R.id) != null)
                {
                    var oldPersistencePart = oldPart.GetPartById((string)axc.Root.Attribute(R.id));

                    var newId2 = "R" + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 16);
                    EmbeddedControlPersistenceBinaryDataPart newPersistencePart = newPart.AddNewPart<EmbeddedControlPersistenceBinaryDataPart>("application/vnd.ms-office.activeX", newId2);

                    newPersistencePart.FeedData(oldPersistencePart.GetStream());
                    axc.Root.Attribute(R.id).Value = newId2;
                    newPart.PutXDocument();
                }
            }
        }

        private static void CopyLegacyDiagramText(OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement textdataReference, XName attributeName)
        {
            string relId = (string)textdataReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;

            var tempPartIdPair = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
            if (tempPartIdPair != default)
                return;

            var oldPart = oldContentPart.GetPartById(relId);

            var newId = "R" + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 16);
            LegacyDiagramTextPart newPart = newContentPart.AddNewPart<LegacyDiagramTextPart>(newId);

            newPart.FeedData(oldPart.GetStream());
            textdataReference.Attribute(attributeName).Value = newId;
        }

        private static void CopyExtendedPart(OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement extendedReference, XName attributeName)
        {
            string relId = (string)extendedReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;
            try
            {
                // First look to see if this relId has already been added to the new document.
                // This is necessary for those parts that get processed with both old and new ids, such as the comments
                // part.  This is not necessary for parts such as the main document part, but this code won't malfunction
                // in that case.
                var tempPartIdPair = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (tempPartIdPair != default)
                    return;

                var tempEr = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
                if (tempEr != null)
                    return;

                var oldPart = (ExtendedPart)oldContentPart.GetPartById(relId);
                var fileInfo = new FileInfo(oldPart.Uri.OriginalString);

                var newPart = newContentPart switch
                {
                    ChartColorStylePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ChartDrawingPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ChartPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ChartsheetPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ChartStylePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CommentAuthorsPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ConnectionsPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ControlPropertiesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CoreFilePropertiesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CustomDataPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CustomDataPropertiesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CustomFilePropertiesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CustomizationPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CustomPropertyPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CustomUIPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CustomXmlMappingsPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CustomXmlPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CustomXmlPropertiesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    DiagramColorsPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    DiagramDataPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    DiagramLayoutDefinitionPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    DiagramPersistLayoutPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    DiagramStylePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    DigitalSignatureOriginPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    DrawingsPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    EmbeddedControlPersistenceBinaryDataPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    EmbeddedControlPersistencePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    EmbeddedObjectPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    EmbeddedPackagePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ExtendedFilePropertiesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ExtendedPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    FontPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    FontTablePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    HandoutMasterPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    InternationalMacroSheetPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    LegacyDiagramTextInfoPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    LegacyDiagramTextPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    MacroSheetPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    NotesMasterPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    NotesSlidePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    PresentationPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    PresentationPropertiesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    RibbonAndBackstageCustomizationsPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    SingleCellTablePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    SlideCommentsPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    SlideLayoutPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    SlideMasterPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    SlidePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    SlideSyncDataPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    StyleDefinitionsPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    StylesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    TableDefinitionPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    TableStylesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ThemeOverridePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ThemePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ThumbnailPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    TimeLineCachePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    TimeLinePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    UserDefinedTagsPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    VbaDataPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    VbaProjectPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ViewPropertiesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    VmlDrawingPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    XmlSignaturePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    _ => null
                };

                relId = newContentPart.GetIdOfPart(newPart);
                using (var stream = oldPart.GetStream())
                    newPart?.FeedData(stream);
                extendedReference.Attribute(attributeName).Value = relId;
            }
            catch (ArgumentOutOfRangeException)
            {
                try
                {
                    var er = oldContentPart.GetExternalRelationship(relId);
                    var newEr = newContentPart.AddExternalRelationship(er.RelationshipType, er.Uri);
                    extendedReference.Attribute(R.id).Value = newEr.Id;
                }
                catch (KeyNotFoundException)
                {
                    var newPart = newContentPart.OpenXmlPackage.GetPackage().GetParts().FirstOrDefault(p => p.Uri == newContentPart.Uri);
                    if (!newPart.Relationships.Contains(relId))
                    {
                        newPart.Relationships.Create(new Uri("NULL", UriKind.RelativeOrAbsolute),
                            System.IO.Packaging.TargetMode.Internal,
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", relId);
                    }
                }
            }
        }

        // General function for handling images that tries to use an existing image if they are the same
        private static ImageData ManageImageCopy(ImagePart oldImage, List<ImageData> images)
        {
            ImageData oldImageData = new ImageData(oldImage);
            foreach (ImageData item in images)
            {
                if (item.Compare(oldImageData))
                    return item;
            }
            images.Add(oldImageData);
            return oldImageData;
        }

        // General function for handling media that tries to use an existing media item if they are the same
        private static MediaData ManageMediaCopy(DataPart oldMedia, List<MediaData> mediaList)
        {
            MediaData oldMediaData = new MediaData(oldMedia);
            foreach (MediaData item in mediaList)
            {
                if (item.Compare(oldMediaData))
                    return item;
            }
            mediaList.Add(oldMediaData);
            return oldMediaData;
        }

        private static void CopyRelatedSound(PresentationDocument newDocument, OpenXmlPart oldContentPart, OpenXmlPart newContentPart,
            XElement soundReference, XName attributeName)
        {
            string relId = (string)soundReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;

            ExternalRelationship alreadyExistingExternalRelationship = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
            if (alreadyExistingExternalRelationship != null)
                return;

            ReferenceRelationship alreadyExistingReferenceRelationship = newContentPart.DataPartReferenceRelationships.FirstOrDefault(er => er.Id == relId);
            if (alreadyExistingReferenceRelationship != null)
                return;

            if (oldContentPart.GetReferenceRelationship(relId) is AudioReferenceRelationship)
            {
                AudioReferenceRelationship temp = (AudioReferenceRelationship)oldContentPart.GetReferenceRelationship(relId);
                MediaDataPart newSound = newDocument.CreateMediaDataPart(temp.DataPart.ContentType);
                newSound.FeedData(temp.DataPart.GetStream());
                AudioReferenceRelationship newRel = null;

                if (newContentPart is SlidePart)
                    newRel = ((SlidePart)newContentPart).AddAudioReferenceRelationship(newSound);
                else if (newContentPart is SlideLayoutPart)
                    newRel = ((SlideLayoutPart)newContentPart).AddAudioReferenceRelationship(newSound);
                else if (newContentPart is SlideMasterPart)
                    newRel = ((SlideMasterPart)newContentPart).AddAudioReferenceRelationship(newSound);
                else if (newContentPart is HandoutMasterPart)
                    newRel = ((HandoutMasterPart)newContentPart).AddAudioReferenceRelationship(newSound);
                else if (newContentPart is NotesMasterPart)
                    newRel = ((NotesMasterPart)newContentPart).AddAudioReferenceRelationship(newSound);
                else if (newContentPart is NotesSlidePart)
                    newRel = ((NotesSlidePart)newContentPart).AddAudioReferenceRelationship(newSound);
                soundReference.Attribute(attributeName).Value = newRel.Id;
            }
            if (oldContentPart.GetReferenceRelationship(relId) is MediaReferenceRelationship)
            {
                MediaReferenceRelationship temp = (MediaReferenceRelationship)oldContentPart.GetReferenceRelationship(relId);
                MediaDataPart newSound = newDocument.CreateMediaDataPart(temp.DataPart.ContentType);
                newSound.FeedData(temp.DataPart.GetStream());
                MediaReferenceRelationship newRel = null;

                if (newContentPart is SlidePart)
                    newRel = ((SlidePart)newContentPart).AddMediaReferenceRelationship(newSound);
                else if (newContentPart is SlideLayoutPart)
                    newRel = ((SlideLayoutPart)newContentPart).AddMediaReferenceRelationship(newSound);
                else if (newContentPart is SlideMasterPart)
                    newRel = ((SlideMasterPart)newContentPart).AddMediaReferenceRelationship(newSound);
                soundReference.Attribute(attributeName).Value = newRel.Id;
            }
        }
    }

    public class PresentationBuilderException : Exception
    {
        public PresentationBuilderException(string message) : base(message) { }
    }

    public class PresentationBuilderInternalException : Exception
    {
        public PresentationBuilderInternalException(string message) : base(message) { }
    }
}
