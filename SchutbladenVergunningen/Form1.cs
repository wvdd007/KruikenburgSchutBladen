using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using LinqToExcel;
using Microsoft.Office.Interop.Word;
using System = Microsoft.Office.Interop.Word.System;
using Windows = Microsoft.Office.Interop.Word.Windows;

namespace SchutbladenVergunningen
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // query naar excel sheet
            var excel = new ExcelQueryFactory("D:\\Users\\alfredo\\Google\\VolleyKrukenburg\\leden 2017-2.xls");
            excel.AddMapping<Lid>(x => x.Naam, "Naam");
            excel.AddMapping<Lid>(x => x.Adres, "Adres");
            excel.AddMapping<Lid>(x => x.Gemeente, "Gemeente");
            excel.AddMapping<Lid>(x => x.Ploeg, "Ploeg");
            excel.AddMapping<Lid>(x => x.Geb, "Geb", s =>
            {
                DateTime result;
                if (DateTime.TryParse(s, CultureInfo.GetCultureInfo("nl-be"), DateTimeStyles.None, out result))
                    return result;
                return null;
            });
            excel.AddMapping<Lid>(x => x.Nat, "Nat");
            excel.AddMapping<Lid>(x => x.Gesl, "Gesl");
            excel.AddMapping<Lid>(x => x.Email, "E-mail");
            excel.AddMapping<Lid>(x => x.Telefoonnummers, "Telefoonnr#", s =>
            {
                return s;
            });
            excel.AddMapping<Lid>(x => x.Licentie, "licentie");
            foreach (var field in excel.GetColumnNames("Blad1"))
            {
                Debug.WriteLine(field);
            }
            var leden = from c in excel.Worksheet<Lid>("Blad1") select c;

            // We lopen over alle leden en steken ze in een lijst per ploeg
            var lijstPerPloeg = new Dictionary<string, List<Lid>>();
            foreach (var lid in leden)
                if (lid.Ploeg != null)
                {
                    var ploegen = lid.Ploeg.Split('-');
                    foreach (var ploeg in ploegen)
                    {
                        List<Lid> ploegLeden;
                        if (!lijstPerPloeg.TryGetValue(ploeg, out ploegLeden))
                        {
                            ploegLeden = new List<Lid>();
                            lijstPerPloeg[ploeg] = ploegLeden;
                        }
                        ploegLeden.Add(lid);
                    }
                }

           

            var doc = new Document();
            var first = true;
            doc.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
            foreach (var ploeg in lijstPerPloeg.Keys.OrderBy(x => x))
            {

                Paragraph par = null;
                if (!first)
                {
                    par=doc.Paragraphs.Add();
                    par.Range.InsertBreak(WdBreakType.wdPageBreak);
                }
                first = false;

                textBox1.AppendText($"Ploeg : {ploeg}\n");
                par = doc.Paragraphs.Add();
                par.set_Style(WdBuiltinStyle.wdStyleHeading1);
                par.Range.Text =($"Ploeg : {ploeg}\n");

                par = doc.Paragraphs.Add();
                par.set_Style(WdBuiltinStyle.wdStyleHeading2); 
                var start = par.Range.Start;

                par.Range.Text=
                    $"Naam\tGeboortejaar\tLicentienummer\tEmail\tTelefoonnummers\n";
                var count = 0;
                foreach (var lid in lijstPerPloeg[ploeg].OrderBy(x => x.Naam))
                {
                    textBox1.AppendText(
                        $"  Lid : {lid.Naam} - {lid.Geb?.Year} - {lid.Licentie} - {lid.Email} - {lid.Telefoonnummers}\n");
                    this.Text = $"{lid.Naam}-{lid.Email}";
                    par = doc.Paragraphs.Add();
                    par.set_Style(WdBuiltinStyle.wdStyleNormal);

                    par.Range.Text=
                        $"{lid.Naam}\t{lid.Geb?.Year}\t{lid.Licentie}\t{lid.Email}\t{lid.Telefoonnummers}\n";
                    count++;
                }
                for(var rest=count ;rest<=20;rest++)
                {
                   
                    par = doc.Paragraphs.Add();
                    par.set_Style(WdBuiltinStyle.wdStyleNormal);

                    par.Range.Text=
                        $"\t\t\t\t\n";
                    count++;
                }
                var end = par.Range.End;
                var table= doc.Range(start, end).ConvertToTable();
                table.ApplyStyleRowBands = true;
                table.ApplyStyleFirstColumn = false;
                table.ApplyStyleColumnBands = false;
                table.ApplyStyleHeadingRows = true;
                table.ApplyStyleColumnBands = false;
                table.ApplyStyleDirectFormatting("Rastertabel 3 - Accent 6");
                table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
            }
            doc.SaveAs($@"d:\doc{DateTime.Now.Ticks}.docx");
            doc.Close();
            this.Close();
        }
    }
}