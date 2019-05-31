using System;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using System.Text.RegularExpressions;

namespace Affirma.OpenSource.PresentationTool
{
    public static class PresentationEngine
    {

        /// <summary>
        /// Generates a new presentation document from the combination of slides from inputted documents. Returns location of generated document.
        /// </summary>
        /// <param name="documentURLs"></param>
        public static string CombinePresentations(IList<string> documentURLs)
        {
            Utils.CheckPresentationFiles(documentURLs);

            string tempPDocLocation = Utils.GetTempFile();

            using (PresentationDocument pDoc = PresentationDocument.Open(tempPDocLocation, true))
            {
                foreach(string pDocFile in documentURLs)
                {
                    using (var tempDoc = PresentationDocument.Open(pDocFile, false)) {
                        var slides = tempDoc.PresentationPart.SlideParts;
                        foreach(var slide in slides)
                        {
                            pDoc.PresentationPart.AddPart(slide);
                        }
                    }
                }
            }

            return tempPDocLocation;
        }

        public static void GenerateSlideFromTemplate(IList<string> textContent, string imageUrl, SlideLayoutPart slideLayout)
        {

        }

        /// <summary>
        /// Builds out image as a Picture shape using BlipFill and appends it to the slide shapetree, using a shape from SlideLayout to determine size and location of image.
        /// </summary>
        /// <param name="slide">The slide to insert the image</param>
        /// <param name="layoutPlaceholderShape">The shape where the image will be inserted</param>
        /// <param name="imageLocation">URL to the local image file</param>
        /// <param name="shapeNumber">A unique positive integer for the shape relationship ID</param>
        public static void AddImageToSlide(this SlidePart slide, Shape layoutPlaceholderShape, string imageLocation, uint shapeNumber)
        {
            ShapeProperties shapeProperties = new ShapeProperties();

            // Generate a unique relationship ID to avoid collision
            string embedId = "rId" + (slide.Slide.Elements().Count() + 100 + shapeNumber).ToString();

            // Determine the image location from layout placeholder shape
            Drawing.Offset imageOffset = (Drawing.Offset)(layoutPlaceholderShape).ShapeProperties.Transform2D.Offset.Clone();
            Drawing.Extents imageExtents = (Drawing.Extents)(layoutPlaceholderShape).ShapeProperties.Transform2D.Extents.Clone();
            Drawing.Transform2D imageTransform = new Drawing.Transform2D();

            imageTransform.Append(imageExtents);
            imageTransform.Offset = imageOffset;
            shapeProperties.Append(imageTransform);

            // Preset geometry
            Drawing.PresetGeometry presetGeometry = new Drawing.PresetGeometry() { Preset = Drawing.ShapeTypeValues.Rectangle };
            Drawing.AdjustValueList adjustValueList = new Drawing.AdjustValueList();

            presetGeometry.Append(adjustValueList);
            shapeProperties.Append(presetGeometry);

            // Generate blip extension
            Drawing.BlipExtensionList blipExtensionList = new Drawing.BlipExtensionList();
            Drawing.BlipExtension blipExtension = new Drawing.BlipExtension() { Uri = $"{{{Guid.NewGuid().ToString()}}}" };

            // Local DPI (dots per square inch)
            DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi useLocalDpi = new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi() { Val = false };
            useLocalDpi.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension.Append(useLocalDpi);
            blipExtensionList.Append(blipExtension);

            // Set up blipfill for image
            Drawing.Blip imageBlip = new Drawing.Blip();
            Drawing.Stretch stretch = new Drawing.Stretch(new Drawing.FillRectangle());
            BlipFill blipFill = new BlipFill();

            imageBlip.Append(blipExtensionList);
            blipFill.Append(imageBlip);
            blipFill.Append(stretch);

            // Set up picture and nonvisual properties
            Picture picture = new Picture();
            NonVisualPictureProperties nonVisualPictureProperties = new NonVisualPictureProperties(
             new NonVisualDrawingProperties() { Id = (50U + shapeNumber), Name = "Picture " + shapeNumber },
             new NonVisualPictureDrawingProperties(new Drawing.PictureLocks() { NoChangeAspect = true }),
             new ApplicationNonVisualDrawingProperties()
             );

            picture.Append(nonVisualPictureProperties);
            picture.Append(blipFill);
            picture.Append(shapeProperties);

            //s.Parent.ReplaceChild(P, s);
            slide.Slide.CommonSlideData.ShapeTree.AppendChild(picture);

            // Determine ImagePart type
            ImagePart imagePart;

            switch(Path.GetExtension(imageLocation))
            {
                case ".jpg":
                case ".jpeg": imagePart = slide.AddImagePart(ImagePartType.Jpeg, embedId); break;
                case ".png": imagePart = slide.AddImagePart(ImagePartType.Png, embedId); break;
                case ".gif": imagePart = slide.AddImagePart(ImagePartType.Gif, embedId); break;
                case ".ico": imagePart = slide.AddImagePart(ImagePartType.Icon, embedId); break;
                default: throw new ArgumentException($"Error: Image file type '{Path.GetExtension(imageLocation)}' for file '{Path.GetFileNameWithoutExtension(imageLocation)}' not recognized!");
            }

            // Read the image file as a stream to ImagePart
            using (FileStream fs = File.OpenRead(imageLocation))
            {
                fs.Seek(0, SeekOrigin.Begin);
                imagePart.FeedData(fs);
            }

            imageBlip.Embed = embedId;
        }

        public static void AddTextToSlide()
        {

        }
    }
}
