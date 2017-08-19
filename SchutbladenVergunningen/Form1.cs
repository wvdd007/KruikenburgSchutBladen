using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Windows.Forms;
using LinqToExcel;
using LinqToExcel.Extensions;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Word;
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
            var excel = new ExcelQueryFactory("D:\\Leden 2017-2 willy.xls");
            excel.AddMapping<Lid>(x => x.Naam, "Naam");
            excel.AddMapping<Lid>(x => x.RugNummer, "Nr", s =>
                {
                    int result;
                    if (int.TryParse(s, out result))
                        return result;
                    return null;
                }
            );
            excel.AddMapping<Lid>(x => x.Adres, "Adres");
            excel.AddMapping<Lid>(x => x.Gemeente, "Postn + Gem#");
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

            //CreateOutlookGroups(lijstPerPloeg);

            CreateWordDocument(lijstPerPloeg);


            this.Close();
        }

        private void CreateOutlookGroups(Dictionary<string, List<Lid>> lijstPerPloeg)
        {
            const string PR_SMTP_ADDRESS =
                "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            var outlook = new Microsoft.Office.Interop.Outlook.Application();
            foreach (var ploeg in lijstPerPloeg.Keys.OrderBy(x => x))
            {
                var distList = outlook.CreateItem(
                        OlItemType.olDistributionListItem)as DistListItem;
                distList.Subject = "2017-" + ploeg;
                var contacts = new List<ContactItem>();
                foreach (var lid in lijstPerPloeg[ploeg])
                {
                    if (string.IsNullOrEmpty(lid.Email))
                    {
                        this.textBox1.AppendText("Geen email voor " + lid.Naam);
                        continue;
                    }
                    var contact = outlook.CreateItem(
                        OlItemType.olContactItem) as ContactItem;
                    //Recipient recip =
                    //    outlook.Session.CreateRecipient(lid.Naam);
                    var pre = lid.Email.Replace(" - ", "|");
                    var splitted = pre.Split('|');
                    foreach (var email in splitted)
                    {
                       var email2 = email.Trim().Replace(",",".");
                        try
                        {
                            var address= new System.Net.Mail.MailAddress(email2);
                        }
                        catch (System.Exception e)
                        {
                            this.textBox1.AppendText("invalid email"+ email2);
                            goto done;
                        }
                        if (lid.Geb.HasValue)
                        {
                            contact.Birthday = lid.Geb.Value;
                        }
                        contact.FullName = "2017-" + lid.Naam;
                        contact.NickName = "2017-" + lid.Naam;
                        contact.Email1Address = email2;
                        contact.BusinessAddress = lid.Adres;
                        contact.BusinessAddressCity = lid.Gemeente;
                        contact.HomeTelephoneNumber = lid.Telefoonnummers;
                        contact.Hobby = "Volleybal";
                        contact.Body = lid.Ploeg;
                        //contact.Display(true);

                        //Resolve the Recipient before calling AddMember//
                        contact.Save();
                        Recipient recip =
                            outlook.Session.CreateRecipient(contact.Email1Address);
                        var result = recip.Resolve();
                        Debug.Assert(result);
                        distList.AddMember(recip);
                        contacts.Add(contact);
                        done:
                        ;
                    }
                }
                distList.Save();
                //distList.Display(true);
                var path = "d:\\"+ ploeg+"2017.msg";
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                distList.SaveAs(path);
                distList.Delete();
                foreach (var contact in contacts)
                {
                    try
                    {
                        contact.Delete();
                    }
                    catch (System.Exception e)
                    {
                        this.textBox1.AppendText("Failed:");
                    }
                }
            }
        }

        private void CreateWordDocument(Dictionary<string, List<Lid>> lijstPerPloeg)
        {
            var doc = new Document();
            var first = true;
            doc.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
            foreach (var ploeg in lijstPerPloeg.Keys.OrderBy(x => x))
            {
                Paragraph par = null;
                if (!first)
                {
                    par = doc.Paragraphs.Add();
                    par.Range.InsertBreak(WdBreakType.wdPageBreak);
                }
                first = false;

                textBox1.AppendText($"Ploeg : {ploeg}\n");
                par = doc.Paragraphs.Add();
                par.set_Style(WdBuiltinStyle.wdStyleHeading1);
                par.Range.Text = ($"Ploeg : {ploeg}\n");

                par = doc.Paragraphs.Add();
                par.set_Style(WdBuiltinStyle.wdStyleHeading2);
                var start = par.Range.Start;

                par.Range.Text =
                    $"Nr\tLicentienummer\tNaam\tGeboortejaar\tEmail\tTelefoonnummers\n";
                var count = 0;
                foreach (var lid in lijstPerPloeg[ploeg]
                    .OrderBy(x => x.RugNummer)
                    .ThenBy(x => x.Naam))
                {
                    textBox1.AppendText(
                        $"  Lid : {lid.RugNummerTekst} - {lid.Naam} - {lid.Geb?.Year} - {lid.Licentie} - {lid.Email} - {lid.Telefoonnummers}\n");
                    this.Text = $"{lid.Naam}-{lid.Email}";
                    par = doc.Paragraphs.Add();
                    par.set_Style(WdBuiltinStyle.wdStyleNormal);

                    par.Range.Text =
                        $"{lid.RugNummerTekst}\t{lid.Licentie}\t{lid.Naam}\t{lid.Geb?.Year}\t{lid.Email}\t{lid.Telefoonnummers}\n";
                    count++;
                }
                for (var rest = count; rest <= 20; rest++)
                {
                    par = doc.Paragraphs.Add();
                    par.set_Style(WdBuiltinStyle.wdStyleNormal);

                    par.Range.Text =
                        $"\t\t\t\t\t\n";
                    count++;
                }
                var end = par.Range.End;
                var table = doc.Range(start, end).ConvertToTable();
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
        }
    }
}