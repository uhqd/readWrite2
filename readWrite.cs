using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using VMS.TPS.Common.Model.API;
using System.Collections.Generic; // for List<>

// pour manipuler les .txt, .csv
using System.IO;
// pour manipuler les .docx
using Microsoft.Office.Interop.Word;
// pour manipuler les .xlsx
using Microsoft.Office.Interop.Excel;
// pour ecrire les pdf
using PdfSharp.Pdf;
using PdfSharp.Drawing;
// pour lire les pdf
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;


[assembly: AssemblyVersion("18.8.0.01")]
namespace VMS.TPS
{
    public class Script
    {
        // 
        public Script()
        {
        }
        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context)
        {
            #region set current dir to the dll dir
            string fullPath = Assembly.GetExecutingAssembly().Location; //get the full location of the assembly          
            string theDirectory = Path.GetDirectoryName(fullPath);//get the folder that's in                                                                  
            Directory.SetCurrentDirectory(theDirectory);// set current directory as the .dll directory...
            #endregion

            #region check if a plan with dose is loaded, no verification plan allowed

            bool aPlanIsLoaded = true;
            try
            {
                string s = context.Patient.Id; // check if a patient is loaded
            }
            catch
            {
                MessageBox.Show("Merci de charger un patient");
                return;
            }

            if (context.PlanSetup == null)
            {

                MessageBox.Show("Aucun plan chargé, les tests de plans et de dose ne seront pas effectués");
                aPlanIsLoaded = false;
            }
            if (aPlanIsLoaded)
            {
                if (context.PlanSetup.PlanIntent == "VERIFICATION")
                {
                    MessageBox.Show("Aucun plan chargé, les tests de plans et de dose ne seront pas effectués");
                    aPlanIsLoaded = false;
                }
                if (!context.PlanSetup.IsDoseValid)
                {
                    MessageBox.Show("Aucune dose assoicée au plan");
                    aPlanIsLoaded = false;
                }
            }

            #endregion


            List<string> mylist = new List<string>();
            foreach (Beam b in context.PlanSetup.Beams)
            {
                mylist.Add(b.Id + ";" + b.EnergyMode + "\n");
            }




            #region ECRIRE UN FICHIER TXT
            string thepath = @"toto.txt";
            File.AppendAllText(thepath, "FX;Energy\n");
            foreach (string s in mylist)
            {
                File.AppendAllText(thepath, s);

            }
            MessageBox.Show("Fichier " + thepath + " bien sauvegardé dans " + Directory.GetCurrentDirectory());

            #endregion

            #region LIRE UN FICHIER TXT

            string[] lines = File.ReadAllLines(thepath);
            string totaltext = string.Empty;
            foreach (string line in lines)
            {
                totaltext += line;
            }

            MessageBox.Show("Fichier " + thepath + " bien lu dans " + Directory.GetCurrentDirectory());
            MessageBox.Show("Voici le contenu\n" + totaltext);

            #endregion

            #region ECRIRE UN FICHIER DOCX
            #region color of the table lines
            var wdcUncheck = (WdColor)(255 + 0x100 * 255 + 0x10000 * 213); // pale yellow
            var wdcX = (WdColor)(252 + 0x100 * 85 + 0x10000 * 62); // pale red
            var wdcWarn = (WdColor)(255 + 0x100 * 188 + 0x10000 * 143); // pale orange
            var wdcInfo = WdColor.wdColorGray05;//pale gray
            var wdcOk = (WdColor)(183 + 0x100 * 255 + 0x10000 * 183); // pale yellow
            #endregion

            #region creation du fichier word et ouverture de word
            Microsoft.Office.Interop.Word.Application winword;
            winword = new Microsoft.Office.Interop.Word.Application();
            winword.ShowAnimation = false;
            winword.Visible = false;
            object missing;
            Microsoft.Office.Interop.Word.Document document;
            missing = System.Reflection.Missing.Value;
            document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            document.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
            #endregion

            #region en tête
            foreach (Microsoft.Office.Interop.Word.Section section in document.Sections)
            {
                Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlue;
                headerRange.Font.Size = 12;
                headerRange.Text = "TEST DE HEADER";
            }
            #endregion

            #region pied de page
            foreach (Microsoft.Office.Interop.Word.Section wordSection in document.Sections)
            {
                //Get the footer range and add the footer details.  
                Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlack;
                footerRange.Font.Size = 10;
                footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                string footText = "TEST DE FOOTER";
                footerRange.Text = footText;// "Footer text goes here";
            }
            #endregion

            #region première ligne simple
            document.Content.SetRange(0, 0);
            // Ajouter du texte
            Paragraph para = document.Paragraphs.Add();
            para.Range.Text = "Julia et Mathilde";
            para.Range.InsertParagraphAfter();
            #endregion

            #region plus compliqué un tableau 
            Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
            para1.Range.Font.Size = 12;
            para1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            para1.Range.Text = Environment.NewLine;
            para1.Range.Text = Environment.NewLine;


            Microsoft.Office.Interop.Word.Table table1 = document.Tables.Add(para1.Range, 4, 4, ref missing, ref missing);
            table1.PreferredWidth = 450.0f;
            table1.Borders.Enable = 1;
            //table1.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent); // Autofit table to content
            table1.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
            foreach (Microsoft.Office.Interop.Word.Row row in table1.Rows)
            {
                foreach (Microsoft.Office.Interop.Word.Cell cell in row.Cells)
                {
                    cell.Range.Font.Bold = 1;
                    cell.Range.Font.Size = 8;

                    var wdc = (WdColor)(229 + 0x100 * 243 + 0x10000 * 229); // pale green
                    cell.Shading.BackgroundPatternColor = wdc;//WdColor.wdColorLightYellow;  // 229 243 229
                    cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    if (cell.ColumnIndex % 2 != 0)
                        cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    else
                        cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                }
            }


            table1.Rows[1].Cells[1].Range.Text = "Patient : ";

            table1.Rows[1].Cells[2].Range.Text = "toto";
            table1.Rows[1].Cells[3].Range.Text = "titi";
            para1.Range.Text = Environment.NewLine;

            #endregion

            #region sauvegarde et fermeture de word
            string textfilename = Directory.GetCurrentDirectory() + @"\toto2.docx";
            object filename = textfilename;
            document.SaveAs2(ref filename);
            document.Close(ref missing, ref missing, ref missing);
            winword.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(document);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(winword);

            #endregion
            MessageBox.Show("Création d'un WORD\nFichier " + textfilename + " créé et sauvegardé");
            #endregion

            #region LIRE UN FICHIER DOCX
            Document doc = null;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            try
            {

                object filePath = textfilename; // Remplace par le bon chemin
                object readOnly = true;
                wordApp.Visible = false;
                doc = wordApp.Documents.Open(ref filePath, ReadOnly: ref readOnly);
                string firstLine = doc.Paragraphs[1].Range.Text.Trim();
                MessageBox.Show("Lecture de la première ligne:\n" + firstLine);
            }
            finally   // sera toujours executé contrairement à catch
            {
                doc?.Close(false);
                wordApp.Quit();
            }
            #endregion

            #region ECRIRE UN FICHIER XLSX 
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                // Créer un nouveau classeur
                workbook = excelApp.Workbooks.Add();
                worksheet = workbook.ActiveSheet;

                // Écrire les valeurs dans les cellules
                worksheet.Cells[1, 1] = 1; // A1
                worksheet.Cells[1, 2] = 2; // B1

                // Définir le chemin du fichier
                string filePath = Directory.GetCurrentDirectory() + @"\toto.xlsx";

                // Sauvegarder le fichier
                workbook.SaveAs(filePath);

            }
            finally
            {
                // Fermeture propre d'Excel
                workbook?.Close(false);
                excelApp.Quit();
            }


            #endregion

            #region LIRE UN FICHIER XLSX 
            Microsoft.Office.Interop.Excel.Application excelApp2 = new Microsoft.Office.Interop.Excel.Application();
            workbook = null;
            worksheet = null;
            try
            {
                // Ouvrir le fichier Excel (modifier le chemin !)
                string filePath = Directory.GetCurrentDirectory() + @"\toto.xlsx";
                workbook = excelApp2.Workbooks.Open(filePath);
                worksheet = workbook.ActiveSheet;

                // Lire la valeur de la cellule B1 (ligne 1, colonne 2)
                object cellValue = worksheet.Cells[1, 2].Value;
                MessageBox.Show("Lecture du fichier EXCEL\nValeur de B1 : " + cellValue);
            }
            finally
            {
                // Fermeture propre d'Excel
                workbook?.Close(false);
                excelApp2.Quit();
            }

            #endregion

            #region Créer un document PDF (avec pdfsharp)
            try
            {
                PdfSharp.Pdf.PdfDocument document2 = new PdfSharp.Pdf.PdfDocument();
                document2.Info.Title = "Mon premier PDF";

                // Ajouter une page
                PdfSharp.Pdf.PdfPage page = document2.AddPage();
                XGraphics gfx = XGraphics.FromPdfPage(page);

                // Définir une police
                XFont font = new XFont("Arial", 20, XFontStyle.Bold);

                // Ajouter du texte
                gfx.DrawString("Mathilde et Julia", font, XBrushes.Black, new XPoint(50, 100));

                // Définir le chemin du fichier
                string filePath2 = Directory.GetCurrentDirectory() + @"\toto.pdf";

                // Sauvegarder le fichier PDF
                document2.Save(filePath2);
                MessageBox.Show("Fichier PDF créé avec succès !");
            }
            catch
            {
                MessageBox.Show("Erreur a la creation du pdf");
            }
            #endregion

            #region LIRE LE PDF (avec itext)
            string filePath3 = Directory.GetCurrentDirectory() + @"\toto.pdf";
            using (PdfReader reader = new PdfReader(filePath3))
            using (iText.Kernel.Pdf.PdfDocument pdfDoc = new iText.Kernel.Pdf.PdfDocument(reader))
            {
                // Extraire tout le texte de la première page
                string text = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(1));

                // Séparer les lignes et afficher la première
                string firstLine = text.Split('\n')[0];
                MessageBox.Show("Première ligne du PDF : " + firstLine);
            }
            #endregion



        }
    }
}

