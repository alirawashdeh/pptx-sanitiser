using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace ConsoleApplication
{
    public class Program
    {
        public static void Main(string[] args)
        {
            String[] filenames = System.IO.Directory.GetFiles(System.IO.Directory.GetCurrentDirectory(),"*.pptx",SearchOption.AllDirectories);
            
            // Process each file
            bool filesProcessed = false;
            for (int i = 0; i < filenames.Length; i++)
            {
                if(filenames[i].Contains("~$") == false)
                {
                    UpdateSlideText(filenames[i]);
                    filesProcessed = true;
                    Console.Out.WriteLine("Updated file: " + filenames[i] + " 👌");
                }
            }

            if (!filesProcessed)
            {
                throw new ArgumentException("No pptx files found.");
            }

        }

    public static void UpdateSlideText(string docName)
    {
        using (PresentationDocument ppt = PresentationDocument.Open(docName, true))
        {
            // Get the relationship ID of the first slide.
            if((ppt.PresentationPart != null) && (ppt.PresentationPart.Presentation.SlideIdList != null))
            {
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

                // Update document properties
                ppt.PackageProperties.Creator = null;
                ppt.PackageProperties.Description = null;
                ppt.PackageProperties.Subject = null;
                ppt.PackageProperties.Category = null;
                ppt.PackageProperties.Keywords = null;
                ppt.PackageProperties.Title = null;
                ppt.PackageProperties.LastModifiedBy = ObfuscateText(ppt.PackageProperties.LastModifiedBy);
                if(ppt.ExtendedFilePropertiesPart != null) 
                {
                    ppt.ExtendedFilePropertiesPart.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties();
                }

                // Update the text on each slide
                for (int i=0; i<part.SlideParts.Count(); i++)
                {
                    SlideId slideId = (SlideId)slideIds[i];
                    String? relId = slideId.RelationshipId;
                    
                    if(slideId.RelationshipId != null)
                    {
                        // Get the slide part from the relationship ID.
                        SlidePart slide = (SlidePart) part.GetPartById(relId);

                        // Get the inner text of the slide:
                        IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();
                        
                        foreach (A.Text text in texts)
                        {
                            Regex pattern = new Regex("[A-Za-z0-9]");
                            text.Text = ObfuscateText(text.Text);
                        }
                    }
                }
            }
        }              
    }   

    public static String ObfuscateText(string text)
    {
        Regex pattern = new Regex("[A-Za-z0-9]");
        return pattern.Replace(text, "x");
    }

    }
}