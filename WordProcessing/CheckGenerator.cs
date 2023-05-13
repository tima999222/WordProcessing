using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PropertiesReflection;

namespace WordProcessing
{
    public class CheckGenerator
    {
        public void GenereateCheck<T>(string pathTemplate, string pathDest, T obj) where T : class
        {
            File.Copy(@"C:\Users\Тимофей\Desktop\check.docx", @"C:\Users\Тимофей\Desktop\check1.docx");

            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(@"C:\Users\Тимофей\Desktop\check1.docx", true))
            {
                var props = obj.GetProperties();

                if (wordprocessingDocument != null)
                {
                    var texts = wordprocessingDocument.MainDocumentPart.Document.Body.Descendants<Text>();

                    foreach (Text text in texts)
                    {
                        foreach (var p in props)
                        {
                            if (("{" + p.Name + "}").Contains(text.Text))
                            {
                                text.Text = p.GetValue(obj).ToString();
                            }
                        }
                    }
                }
            }
        }
    }
}
