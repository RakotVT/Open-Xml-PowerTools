﻿// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

/***************************************************************************

Copyright (c) Microsoft Corporation 2010.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license
can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

***************************************************************************/

/***************************************************************************
 * IMPORTANT NOTE:
 * 
 * With versions 4.1 and later, the name of the HtmlConverter class has been
 * changed to WmlToHtmlConverter, to make it orthogonal with HtmlToWmlConverter.
 * 
 * There are thin wrapper classes, HtmlConverter, and HtmlConverterSettings,
 * which maintain backwards compat for code that uses the old name.
 * 
 * Other than the name change of the classes themselves, the functionality
 * in WmlToHtmlConverter is identical to the old HtmlConverter class.
***************************************************************************/

using System;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using SkiaSharp;

class HtmlConverterHelper
{
    static void Main(string[] args)
    {
        var n = DateTime.Now;
        var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
        tempDi.Create();

        /*
         * This example loads each document into a byte array, then into a memory stream, so that the document can be opened for writing without
         * modifying the source document.
         */
        foreach (var file in Directory.GetFiles("../../", "*.docx"))
        {
            ConvertToHtml(file, tempDi.FullName);
        }
    }

    public static void ConvertToHtml(string file, string outputDirectory)
    {
        var fi = new FileInfo(file);
        Console.WriteLine(fi.Name);
        byte[] byteArray = File.ReadAllBytes(fi.FullName);
        using (MemoryStream memoryStream = new MemoryStream())
        {
            memoryStream.Write(byteArray, 0, byteArray.Length);
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(memoryStream, true))
            {
                var destFileName = new FileInfo(fi.Name.Replace(".docx", ".html"));
                if (outputDirectory != null && outputDirectory != string.Empty)
                {
                    DirectoryInfo di = new DirectoryInfo(outputDirectory);
                    if (!di.Exists)
                    {
                        throw new OpenXmlPowerToolsException("Output directory does not exist");
                    }
                    destFileName = new FileInfo(Path.Combine(di.FullName, destFileName.Name));
                }
                var imageDirectoryName = destFileName.FullName.Substring(0, destFileName.FullName.Length - 5) + "_files";
                int imageCounter = 0;

                var pageTitle = fi.FullName;
                var part = wDoc.CoreFilePropertiesPart;
                if (part != null)
                {
                    pageTitle = (string) part.GetXDocument().Descendants(DC.title).FirstOrDefault() ?? fi.FullName;
                }

                // TODO: Determine max-width from size of content area.
                HtmlConverterSettings settings = new HtmlConverterSettings()
                {
                    AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
                    PageTitle = pageTitle,
                    FabricateCssClasses = true,
                    CssClassPrefix = "pt-",
                    RestrictToSupportedLanguages = false,
                    RestrictToSupportedNumberingFormats = false,
                    ImageHandler = imageInfo =>
                    {
                        DirectoryInfo localDirInfo = new DirectoryInfo(imageDirectoryName);
                        if (!localDirInfo.Exists)
                            localDirInfo.Create();
                        ++imageCounter;
                        string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                        SKEncodedImageFormat imageFormat = (SKEncodedImageFormat)(-1);
                        if (extension == "png")
                            imageFormat = SKEncodedImageFormat.Png;
                        else if (extension == "gif")
                            imageFormat = SKEncodedImageFormat.Gif;
                        else if (extension == "bmp")
                            imageFormat = SKEncodedImageFormat.Bmp;
                        else if (extension == "jpeg")
                            imageFormat = SKEncodedImageFormat.Jpeg;
                        else if (extension == "tiff")
                        {
                            // Convert tiff to gif.
                            extension = "gif";
                            imageFormat = SKEncodedImageFormat.Gif;
                        }

                        // If the image format isn't one that we expect, ignore it,
                        // and don't return markup for the link.
                        if (imageFormat < 0)
                            return null;

                        string imageFileName = imageDirectoryName + "/image" +
                            imageCounter.ToString() + "." + extension;
                        try
                        {
                            using (var imageFile = File.Create(imageFileName))
                                imageInfo.Bitmap.Encode(imageFile, imageFormat, 100);
                        }
                        catch (ExternalException)
                        {
                            return null;
                        }
                        string imageSource = localDirInfo.Name + "/image" +
                            imageCounter.ToString() + "." + extension;

                        XElement img = new XElement(Xhtml.img,
                            new XAttribute(NoNamespace.src, imageSource),
                            imageInfo.ImgStyleAttribute,
                            imageInfo.AltText != null ?
                                new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                        return img;
                    }
                };
                XElement htmlElement = HtmlConverter.ConvertToHtml(wDoc, settings);

                // Produce HTML document with <!DOCTYPE html > declaration to tell the browser
                // we are using HTML5.
                var html = new XDocument(
                    new XDocumentType("html", null, null, null),
                    htmlElement);

                // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
                // XEntity.  PtOpenXmlUtil.cs define the XEntity class.  See
                // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
                // for detailed explanation.
                //
                // If you further transform the XML tree returned by ConvertToHtmlTransform, you
                // must do it correctly, or entities will not be serialized properly.

                var htmlString = html.ToString(SaveOptions.DisableFormatting);
                File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
            }
        }
    }
}
