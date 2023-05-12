using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WorkWith
{


    class Program
    {
        [Obsolete]
        static void Main(string[] args)
        {
            string templateFilePath = "C:\\Users\\Anton\\Desktop\\UMl.docx";
            string outputFilePath = "C:\\Users\\Anton\\Desktop\\My.docx";

            Dictionary<string, string> tagValues = new Dictionary<string, string>();
            tagValues.Add("[NAME]", "АнтонПрилепин");
            //tagValues.Add("[DATE]", DateTime.Now.ToShortDateString());
            //tagValues.Add("[DOM]", "Москва");

            CreateWordDocumentFromTemplate(templateFilePath, outputFilePath, tagValues);

            Console.WriteLine("Документ успешно создан.");
            Console.ReadLine();
        }

        [Obsolete]
        public static void CreateWordDocumentFromTemplate(string templateFilePath, string outputFilePath, Dictionary<string, string> tagValues)
        {
            // Открываю шаблон документа
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(templateFilePath, true))
            {
                // Получаю корневой элемент документа
                Body body = wordDoc.MainDocumentPart.Document.Body;

                // Заменяю тегов в документе на значения из списка
                foreach (var tagValue in tagValues)
                {
                    var tag = new Text(tagValue.Value);
                    var tagElements = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Where(t => t.Text == tagValue.Key).ToList();
                    foreach (var tagElement in tagElements)
                    {
                        tagElement.Text = "";
                        tagElement.InsertAfterSelf(tag);
                        tagElement.Remove();
                    }
                }

                // Сохранение изменений в документе
                wordDoc.MainDocumentPart.Document.Save();

                wordDoc.SaveAs(outputFilePath);
                wordDoc.Close();
               // System.IO.File.Copy(templateFilePath, outputFilePath, true);
            }
        }
    }
}
