using System;
using Spire.Doc;
using Spire.Doc.Documents;

namespace Convertidor_HTML_Word
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Creacion de instancia del documento HTML
            Document document = new Document();

            //Carga del documento a la instancia usando URL fisica del HTML
            document.LoadFromFile("C:\\Users\\Steven\\Documents\\Convertidor\\Doc\\proformaword.html");


            //Guarda el documento en formato word en una ruta fisica 
            document.SaveToFile("C:\\Users\\Steven\\Documents\\Convertidor\\Doc\\HtmltoWord.docx", FileFormat.Docx2013);
        }
    }
}