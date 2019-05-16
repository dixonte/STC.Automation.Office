using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = STC.Automation.Office.Excel;
using Word = STC.Automation.Office.Word;
using ADODB = STC.Automation.Office.ADODB;
using System.Reflection;
using System.IO;
using System.Data.SqlClient;
using STC.Automation.Office.Core;
using STC.Automation.Office.Excel.Utilities;
using STC.Automation.Office.Outlook;

namespace Tester
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnNewExcel_Click(object sender, EventArgs e)
        {
            ADODB.Recordset rs = new ADODB.Recordset();
            rs.Fields.Append("A", STC.Automation.Office.ADODB.Enums.DataType.VarWChar, 100);
            rs.Fields.Append("B", STC.Automation.Office.ADODB.Enums.DataType.VarWChar, 100);
            rs.Fields.Append("C", STC.Automation.Office.ADODB.Enums.DataType.VarWChar, 100);
            rs.Open();
            rs.AddNew(new string[] { "A", "B", "C" }, new object[] { "A1", "B2", "C3" });
            rs.AddNew(new string[] { "A", "B", "C" }, new object[] { "A4", "B5", "C6" });
            rs.AddNew(new string[] { "A", "B", "C" }, new object[] { "A7", "B8", "C9" });

            // Excel
            using (var excel = new Excel.Application())
            {
                excel.NewWorkbook += new STC.Automation.Office.Excel.Events.NewWorkbookEventHandler(excel_NewWorkbook);

                excel.Visible = true;
                
                MessageBox.Show("Version: " + excel.Version.ToString());
                
                using (var workbook = (sender == btnNewExcel) ?
                    excel.Workbooks.Add() :
                    excel.Workbooks.Open(Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName) + "\\Templates\\Open Tester.xls"))
                {
                    using (var worksheet = workbook.ActiveSheet)
                    {
                        using (var range = worksheet.Cells)
                        {
                            (range[1, 1]).Value = "Test";
                        }

                        using (var range = worksheet.Range("A2"))
                        {
                            range.CopyFromRecordset(rs.InternalObject, null, null);
                        }

                        using (var range = worksheet.Range("A2:B3"))
                        {
                            range.Font.Bold = true;
                            range.Font.Color = Color.Teal;
                            range.Font.Italic = true;
                        }

                        using (var range = worksheet.Range("D2:G10"))
                        {
                            string imgPath = Path.Combine(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), @"Resources\Koala.jpg");
                            Image img = null;
                            try
                            {
                                img = Image.FromFile(imgPath);
                            }
                            catch { }
                            //worksheet.Shapes.AddPicture(img, range, true, true).Dispose();

                            if (img != null)
                            {
                                using (var shape = worksheet.Shapes.AddPicture(imgPath, false, true, range, true))
                                {
                                    MessageBox.Show(shape.Name);

                                    worksheet.Hyperlinks.Add(shape, "http://www.google.com/").Dispose();
                                }
                            }
                        }

                        using (var range = worksheet.Range("B2"))
                        {
                            range.AddComment("A comment on cell B2");
                            worksheet.Hyperlinks.Add(range, "http://lmgtfy.com/?q=excel+automation", screenTip: "Let Me Google That For You", textToDisplay: "LMGTFY").Dispose();
                        }

                        using (var range = worksheet.Range("B2"))
                        {
                            if (range.Comment != null)
                                range.Comment.Text(" - New text", 20, false);
                        }

                        using (var range = worksheet.Range("B3"))
                        {
                            if (range.Comment == null)
                                range.AddComment("Another text");

                            if (range.Comment == null)
                                range.AddComment("This should never be seen");
                        }

                        using (var range = worksheet.Range("C2"))
                        {
                            using (var interior = range.Interior)
                            {
                                interior.Color = Color.IndianRed;
                                range.AddComment(interior.Color.ToString());
                            }
                        }
                    }

                    using (var testWorksheet = workbook.Worksheets.Add() as STC.Automation.Office.Excel.Worksheet)
                    {
                        testWorksheet.Name = "Programmatic Worksheet";
                        using (var range = testWorksheet.Cells)
                        {
                            range[1, 1].Value = "Worksheet #2";
                        }

                        using (var chart = workbook.Sheets.Add(testWorksheet, type: Excel.Enums.SheetType.Chart) as STC.Automation.Office.Excel.Chart)
                        {
                            chart.Name = "Programmatic Chart";
                        }
                    }

                    //workbook.Close();
                }

                //excel.Quit();
            }

            //Microsoft.Office.Interop.Excel.ApplicationClass excel = new Microsoft.Office.Interop.Excel.ApplicationClass();

            //excel.Visible = true;

            //var workbook = excel.Workbooks.Add(DBNull.Value);
            //Microsoft.Office.Interop.Excel.Worksheet sheet = workbook.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

            //Microsoft.Office.Interop.Excel.Range range = sheet.Cells[2, 1] as Microsoft.Office.Interop.Excel.Range;
            //range.set_Item(0, 1, "Test");

            //range.CopyFromRecordset(rs.InternalObject, System.Reflection.Missing.Value, 10);
        }

        private void ExcelToPdfButton_Click(object sender, EventArgs e)
        {
            // Copy and paste job from btnNewExcel_Click
            ADODB.Recordset rs = new ADODB.Recordset();
            rs.Fields.Append("A", STC.Automation.Office.ADODB.Enums.DataType.VarWChar, 100);
            rs.Fields.Append("B", STC.Automation.Office.ADODB.Enums.DataType.VarWChar, 100);
            rs.Fields.Append("C", STC.Automation.Office.ADODB.Enums.DataType.VarWChar, 100);
            rs.Open();
            rs.AddNew(new string[] { "A", "B", "C" }, new object[] { "A1", "B2", "C3" });
            rs.AddNew(new string[] { "A", "B", "C" }, new object[] { "A4", "B5", "C6" });
            rs.AddNew(new string[] { "A", "B", "C" }, new object[] { "A7", "B8", "C9" });

            // Excel
            using (var excel = new Excel.Application())
            {
                excel.NewWorkbook += new STC.Automation.Office.Excel.Events.NewWorkbookEventHandler(excel_NewWorkbook);

                excel.Visible = true;
                excel.ScreenUpdating = false;
                excel.DisplayAlerts = false;

                MessageBox.Show("Version: " + excel.Version.ToString());

                using (var workbook = (sender == btnNewExcel) ?
                    excel.Workbooks.Add() :
                    excel.Workbooks.Open(Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName) + "\\Templates\\Open Tester.xls"))
                {
                    using (var worksheet = workbook.ActiveSheet)
                    {
                        using (var range = worksheet.Cells)
                        {
                            (range[1, 1]).Value = "Test";
                        }

                        using (var range = worksheet.Range("A2"))
                        {
                            range.CopyFromRecordset(rs.InternalObject, null, null);
                        }

                        using (var range = worksheet.Range("A2:B3"))
                        {
                            range.Font.Bold = true;
                            range.Font.Color = Color.Teal;
                            range.Font.Italic = true;
                        }

                        using (var range = worksheet.Range("D2:G10"))
                        {
                            string imgPath = Path.Combine(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), @"Resources\Koala.jpg");
                            Image img = null;
                            try
                            {
                                img = Image.FromFile(imgPath);
                            }
                            catch { }
                            //worksheet.Shapes.AddPicture(img, range, true, true).Dispose();

                            if (img != null)
                            {
                                using (var shape = worksheet.Shapes.AddPicture(imgPath, false, true, range, true))
                                {
                                    MessageBox.Show(shape.Name);

                                    worksheet.Hyperlinks.Add(shape, "http://www.google.com/").Dispose();
                                }
                            }
                        }

                        using (var range = worksheet.Range("B2"))
                        {
                            range.AddComment("A comment on cell B2");
                            worksheet.Hyperlinks.Add(range, "http://lmgtfy.com/?q=excel+automation", screenTip: "Let Me Google That For You", textToDisplay: "LMGTFY").Dispose();
                        }

                        using (var range = worksheet.Range("B2"))
                        {
                            if (range.Comment != null)
                                range.Comment.Text(" - New text", 20, false);
                        }

                        using (var range = worksheet.Range("B3"))
                        {
                            if (range.Comment == null)
                                range.AddComment("Another text");

                            if (range.Comment == null)
                                range.AddComment("This should never be seen");
                        }

                        using (var range = worksheet.Range("C2"))
                        {
                            using (var interior = range.Interior)
                            {
                                interior.Color = Color.IndianRed;
                                range.AddComment(interior.Color.ToString());
                            }
                        }
                    }

                    using (var testWorksheet = workbook.Worksheets.Add() as STC.Automation.Office.Excel.Worksheet)
                    {
                        testWorksheet.Name = "Programmatic Worksheet";
                        using (var range = testWorksheet.Cells)
                        {
                            range[1, 1].Value = "Worksheet #2";
                        }

                        using (var chart = workbook.Sheets.Add(testWorksheet, type: Excel.Enums.SheetType.Chart) as STC.Automation.Office.Excel.Chart)
                        {
                            chart.Name = "Programmatic Chart";
                        }
                    }

                    // Code stolen from http://stackoverflow.com/a/7401831/23401
                    var exportSuccessful = true;
                    var outputPath = $@"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\test.pdf";
                    try
                    {
                        // Call Excel's native export function (valid in Office 2007 and Office 2010, AFAIK)
                        workbook.ExportAsFixedFormat(Excel.Enums.FixedFormatType.TypePDF, outputPath);
                    }
                    catch (System.Exception ex)
                    {
                        // Mark the export as failed for the return value...
                        exportSuccessful = false;

                        // Do something with any exceptions here, if you wish...
                        // MessageBox.Show...        
                    }

                    // You can use the following method to automatically open the PDF after export if you wish
                    // Make sure that the file actually exists first...
                    if (System.IO.File.Exists(outputPath))
                    {
                        System.Diagnostics.Process.Start(outputPath);
                    }

                    if (!exportSuccessful )
                    {
                        MessageBox.Show("Uh oh. Something went wrong", "Excel to pdf");
                    }

                    //workbook.Close();
                }

                //excel.Quit();
            }

            //Microsoft.Office.Interop.Excel.ApplicationClass excel = new Microsoft.Office.Interop.Excel.ApplicationClass();

            //excel.Visible = true;

            //var workbook = excel.Workbooks.Add(DBNull.Value);
            //Microsoft.Office.Interop.Excel.Worksheet sheet = workbook.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

            //Microsoft.Office.Interop.Excel.Range range = sheet.Cells[2, 1] as Microsoft.Office.Interop.Excel.Range;
            //range.set_Item(0, 1, "Test");

            //range.CopyFromRecordset(rs.InternalObject, System.Reflection.Missing.Value, 10);
        }

        void excel_NewWorkbook(STC.Automation.Office.Excel.Workbook workbook)
        {
            
        }

        private void btnExcelByProcess_Click(object sender, EventArgs e)
        {
            ADODB.Recordset rs = new ADODB.Recordset();
            rs.Fields.Append("A", STC.Automation.Office.ADODB.Enums.DataType.VarWChar, 100);
            rs.Fields.Append("B", STC.Automation.Office.ADODB.Enums.DataType.VarWChar, 100);
            rs.Fields.Append("C", STC.Automation.Office.ADODB.Enums.DataType.VarWChar, 100);
            rs.Open();
            rs.AddNew(new string[] { "A", "B", "C" }, new object[] { "A1", "B2", "C3" });
            rs.AddNew(new string[] { "A", "B", "C" }, new object[] { "A4", "B5", "C6" });
            rs.AddNew(new string[] { "A", "B", "C" }, new object[] { "A7", "B8", "C9" });

            // Excel
            using (var excel = Excel.Application.FromProcess(System.Diagnostics.Process.GetProcessesByName("EXCEL")[0]))
            {
                if (excel == null)
                {
                    MessageBox.Show("No existing running instance of Excel found.");
                    return;
                }

                excel.Visible = true;

                using (var workbook = excel.Workbooks.Add())
                {
                    using (var worksheet = workbook.ActiveSheet)
                    {
                        using (var range = worksheet.Cells)
                        {
                            range[1, 1].Value = "Test";
                        }

                        using (var range = worksheet.Range("A2"))
                        {
                            range.CopyFromRecordset(rs.InternalObject, null, null);
                        }

                        using (var range = worksheet.Range("D2:G10"))
                        {
                            Image img = Image.FromFile(@"c:\users\tdixon\pictures\grid1.png");

                            //worksheet.Shapes.AddPicture(img, range, true, true).Dispose();

                            /*using (var shape = worksheet.Shapes.AddPicture(@"c:\users\tdixon\pictures\grid1.png", false, true, range, true))
                            {
                                MessageBox.Show(shape.Name);
                            }*/
                        }
                    }

                    workbook.Close();
                }

                excel.Quit();
            }
        }

        private void btnExistingWord_Click(object sender, EventArgs e)
        {
            using (var word = new Word.Application())
            {
                word.Visible = true;

                Console.WriteLine("Version: " + word.Version.ToString());

                //word.Documents.Add().Dispose();

                using (var doc = (sender == btnNewWord) ?
                    word.Documents.Add() :
                    word.Documents.Open(Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName) + "\\Templates\\Open Tester.doc"))
                {
                    //MessageBox.Show(doc.Name);

                    //doc.SaveAs("c:\\test.doc");

                    doc.Activate();

                    Console.WriteLine(doc.Name);
                    Console.WriteLine(doc.FullName);
                    Console.WriteLine(doc.Path);

                    Console.WriteLine("Read only?: " + doc.ReadOnly.ToString());

                    //doc.Close(false);

                    //MessageBox.Show(doc.Sections.Count.ToString());
                    //int x = 0;
                    foreach (Word.Section section in doc.Sections)
                    {
                        //x++;
                        //MessageBox.Show("Index: " + section.Index);
                        //MessageBox.Show("Section: " + section.Index + "\n\n" + section.Range.Text);
                    }
                    //MessageBox.Show(x.ToString());

                    using (Word.Section sec1 = doc.Sections[1])
                    {
                        //MessageBox.Show("Section 1 index: " + sec1.Index.ToString());

                        using (Word.HeadersFooters headers = sec1.Headers)
                        {
                            //MessageBox.Show("Headers: " + headers.Count.ToString());
                            int x = 0;
                            foreach (Word.HeaderFooter header in headers)
                            {
                                x++;

                                //MessageBox.Show("Index: " + header.Index + (header.IsHeader ? " is header" : " is footer") + "\n\n" + header.Range.Text);

                                if (x == 1)
                                    header.Range.Text = "Some other header";
                            }
                            //MessageBox.Show(x.ToString());
                        }

                        using (Word.HeadersFooters footers = sec1.Footers)
                        {
                            //MessageBox.Show("Footers: " + footers.Count.ToString());
                            int x = 0;
                            foreach (Word.HeaderFooter footer in footers)
                            {
                                x++;
                                //MessageBox.Show("Index: " + footer.Index + (footer.IsHeader ? " is header" : " is footer") + "\n\n" + footer.Range.Text);
                            }
                            //MessageBox.Show(x.ToString());
                        }
                    }

                    using (var w = doc.ActiveWindow)
                        w.WindowState = STC.Automation.Office.Word.Enums.WindowState.Normal;

                    if (doc.Bookmarks.Exists("text"))
                    {
                        Console.WriteLine("Found 'text' bookmark");

                        using (var bookmark = doc.Bookmarks["text"])
                        {
                            bookmark.Select();
                        }
                    }

                    if (doc.Bookmarks.Exists("Red"))
                    {
                        Console.WriteLine("Found 'Red' bookmark");

                        doc.Bookmarks["Red"].Select();

                        using (var font = word.Selection.Font)
                        {
                            Console.WriteLine(font.Color.ToString());
                            font.Color = Color.Teal;
                        }
                    }

                    if (doc.Bookmarks.Exists("Bold"))
                    {
                        doc.Bookmarks["Bold"].Select();
                        using (var font = word.Selection.Font)
                        {
                            Console.WriteLine("Bold is bold: {0}", font.Bold);
                            font.Bold = null;
                        }
                    }

                    if (doc.Bookmarks.Exists("NotBold"))
                    {
                        doc.Bookmarks["NotBold"].Select();
                        using (var font = word.Selection.Font)
                        {
                            Console.WriteLine("NotBold is bold: {0}", font.Bold);
                            font.Bold = null;
                        }
                    }

                    if (doc.Bookmarks.Exists("PartlyBold"))
                    {
                        doc.Bookmarks["PartlyBold"].Select();
                        using (var font = word.Selection.Font)
                        {
                            Console.WriteLine("PartlyBold is bold: {0}", font.Bold);
                            font.Bold = null;
                        }
                    }

                    if (doc.Bookmarks.Exists("TwoColumn"))
                    {
                        doc.Bookmarks["TwoColumn"].Select();
                        using (var cols = word.Selection.Columns)
                            Console.WriteLine("TwoColumn column count: {0}", cols.Count);
                    }

                    if (doc.Bookmarks.Exists("ThreeColumn"))
                    {
                        doc.Bookmarks["ThreeColumn"].Select();
                        using (var cols = word.Selection.Columns)
                            Console.WriteLine("ThreeColumn column count: {0}", cols.Count);
                    }

                    doc.Select();
                    using (var find = word.Selection.Find)
                    {
                        find.ClearFormatting();
                        find.Text = "Some text";
                        find.Replacement.ClearFormatting();
                        find.Replacement.Text = "Other bologna";
                        find.Execute(STC.Automation.Office.Word.Enums.Replace.One);
                    }
                }
                

                /*foreach (Word.Document doc in word.Documents)
                {
                    MessageBox.Show(doc.Name);

                    doc.Dispose();

                    break;
                }*/

                //MessageBox.Show("Closed");

                //word.Quit();
            }
        }

        private void btnWordByProc_Click(object sender, EventArgs e)
        {
            using (var filter = new MessageFilter.SimpleMessageFilter())
            {

                filter.CalleeBusy += new MessageFilter.CalleeBusyHandler(filter_CalleeBusy);

                using (var word = Word.Application.FromProcess(System.Diagnostics.Process.GetProcessesByName("WINWORD")[0]))
                {
                    word.Visible = true;

                    /*using (var doc = word.Documents.Add())
                    {
                        doc.SaveAs(Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName) + "\\Empty.doc", STC.Automation.Office.Word.Enums.SaveFormat.Document);
                    }*/

                    foreach (Word.Document doc in word.Documents)
                    {
                        MessageBox.Show(doc.Name);
                    }
                }
            }
        }

        void filter_CalleeBusy(object sender, MessageFilter.CalleeBusyEventArgs args)
        {
            if (args.TickCount > 2000)
            {
                var res = MessageFilter.OleUiBusyDialog.Show(this, args.TaskHandle, "Test");

                if (res == MessageFilter.OleUiBusyDialog.OLEUIFlags.OLEUI_CANCEL)
                {
                    args.Cancel = true;
                }
            }

            args.RetryDelay = 500;
        }

        private void btnFromDataTable_Click(object sender, EventArgs e)
        {
            using (SqlConnection cnn = new SqlConnection(@"Server=lau-sql2005\development;Database=PRImateDB;Trusted_Connection=True;"))
            {
                try
                {
                    cnn.Open();
                }
                catch
                {
                    MessageBox.Show("connection failed", "failure", MessageBoxButtons.OK);
                    return;
                }

                DataTable dt = new DataTable() ;
                SqlDataAdapter da;
                SqlCommand cmd = new SqlCommand("select * from viewRPT_DIER_ProposalBody where ProposalID = -613377248",cnn);

                da = new SqlDataAdapter(cmd);
                da.Fill(dt);

                ADODB.Recordset rs = ADODB.Recordset.FromDataTable(dt);
            }

            //using (SqlConnection cnn = new SqlConnection("Server=lau-sql2005;Database=PRImateDB;Trusted_Connection=True;"))
            //{
            //    try
            //    {
            //        cnn.Open();
            //    }
            //    catch
            //    {
            //        MessageBox.Show("connection failed", "failure", MessageBoxButtons.OK);
            //        return;
            //    }

            //    DataTable dt = new DataTable();
            //    SqlDataAdapter da;
            //    SqlCommand cmd = new SqlCommand("select * from listEmployees", cnn);

            //    da = new SqlDataAdapter(cmd);
            //    da.Fill(dt);

            //    ADODB.Recordset rs = ADODB.Recordset.FromDataTable(dt);
            //}
        }

        private void btnNewAccess_Click(object sender, EventArgs e)
        {
            using (var access = new STC.Automation.Office.Access.Application())
            {
                access.Visible = true;

                MessageBox.Show(access.Version.ToString());
            }
        }

        private void btnExistingAccess_Click(object sender, EventArgs e)
        {
            cbxControlBars.Items.Clear();

            using (var access = STC.Automation.Office.Access.Application.FromProcess(System.Diagnostics.Process.GetProcessesByName("msaccess")[0]))
            {
                foreach (var bar in access.CommandBars)
                {
                    cbxControlBars.Items.Add(bar.Name);
                }
            }
        }

        private void AddChildren(ToolStripMenuItem tsmiParent, CommandBarPopup cbpParent)
        {
            foreach (var control in cbpParent.Controls)
            {
                if (control is CommandBarComboBox)
                {
                    // TODO
                }
                else
                {
                    var newItem = new ToolStripMenuItem(control.Caption);

                    if (control is CommandBarButton)
                    {
                        //newItem.Image = ((CommandBarButton)control).Picture;
                        newItem.Image = newItem.Image = ((CommandBarButton)control).Picture;

                        var but = (CommandBarButton)control;

                        string onAction = "OnAction: " + but.OnAction + "\rParameter: " + but.Parameter;

                        newItem.Click += new EventHandler((s, ev) => { MessageBox.Show(onAction); });
                    }

                    tsmiParent.DropDownItems.Add(newItem);

                    if (control is CommandBarPopup)
                    {
                        AddChildren(newItem, (CommandBarPopup)control);
                    }
                }

                control.Dispose();
            }
        }

        private void cbxControlBars_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(cbxControlBars.SelectedItem.ToString()))
                return;

            using (var access = STC.Automation.Office.Access.Application.FromProcess(System.Diagnostics.Process.GetProcessesByName("msaccess")[0]))
            {
                //MessageBox.Show((access.CommandBars["MISTimeSheet"].Controls["Administration"] as STC.Automation.Office.Core.CommandBarPopup).Controls["Developer Privilege"].Caption);

                toolStrip.Items.Clear();

                using (var bar = access.CommandBars[cbxControlBars.SelectedItem.ToString()])
                {
                    foreach (var control in bar.Controls)
                    {
                        if (control is CommandBarComboBox)
                        {
                            // TODO
                        }
                        else
                        {
                            var newItem = new ToolStripMenuItem(control.Caption);

                            if (control is CommandBarButton)
                            {
                                newItem.Image = ((CommandBarButton)control).Picture;

                                var but = (CommandBarButton)control;

                                string onAction = "OnAction: " + but.OnAction + "\rParameter: " + but.Parameter;

                                newItem.Click += new EventHandler((s, ev) => { MessageBox.Show(onAction); });
                            }

                            toolStrip.Items.Add(newItem);

                            if (control is CommandBarPopup)
                            {
                                AddChildren(newItem, (CommandBarPopup)control);
                            }
                        }

                        control.Dispose();
                    }
                }
            }
        }

        private void btnAutoFilter_Click(object sender, EventArgs e)
        {
            int colCount = 32;
            int dataCount = 1648;
            int lastRow = 1651;

            using (var excel = new Excel.Application())
            {
                excel.Visible = true;

                using (var workbook = excel.Workbooks.Open(Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName) + "\\Templates\\AutoFilter.xls"))
                {
                    using (var worksheet = workbook.ActiveSheet)
                    {
                        worksheet.AutoFilterMode = false;

                        int start = 0;
                        using (var r = worksheet.Range("__Table__"))
                            start = r.Row;

                        using (var r = worksheet.Range(STC.Automation.Office.Excel.Utilities.Ranges.Format(start - 1, 1)))
                            r.AutoFilter();

                        for (int c = 1; c <= colCount; c++)
                        {
                            using (var r = worksheet.Range(STC.Automation.Office.Excel.Utilities.Ranges.Format(start, c, start + dataCount, c)))
                            {
                                // set formats before calculating column autofit
                                if (c >= 13 && c <= 20)
                                    r.NumberFormat = Ranges.ConvertFormat("C0");

                                if (c >= 21 && c <= 24)
                                    r.NumberFormat = Ranges.ConvertFormat("dd/MM/yyyy");


                                r.Columns.AutoFit();
                                if (r.ColumnWidth > 80)
                                    r.ColumnWidth = 80;
                            }
                        }
                       
                        //// sort the AutoFilter by data in column B
                        //using (var filter = worksheet.AutoFilter)
                        //{
                        //    filter.Sort.SortFields.Clear();

                        //    using (var range = worksheet.Range("B4:B" + lastRow))
                        //        filter.Sort.SortFields.Add(range, Excel.Enums.SortOn.Values, Excel.Enums.SortOrder.Ascending, Excel.Enums.SortDataOption.Normal);

                        //    filter.Sort.Header = Excel.Enums.YesNoGuess.Yes;
                        //    filter.Sort.MatchCase = false;
                        //    filter.Sort.Orientation = Excel.Enums.SortOrientation.Columns;
                        //    filter.Sort.Apply();
                        //}
                    }

                    excel.DisplayAlerts = false;
                    workbook.Save();
                    excel.DisplayAlerts = true;
                    //workbook.Close();
                }

                //excel.Quit();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (var word = new Word.Application())
            {
                word.Visible = true;

                Console.WriteLine("Version: " + word.Version.ToString());

                //word.Documents.Add().Dispose();

                using (var doc = word.Documents.Add())
                {
                    word.Selection.TypeText("fred");

                    doc.ExportAsFixedFormat(@"c:\temp\out.pdf", Word.Enums.ExportFormat.PDF);

                }
            }
        }

        private void btnExcelListOpen_Click(object sender, EventArgs e)
        {
            //var apps = Excel.Application.GetRunningApplications();

            //var sb = new StringBuilder();

            //foreach (var app in apps)
            //{
            //    foreach (var workbook in app.Workbooks)
            //    {
            //        sb.AppendLine(workbook.FullName);
            //    }

            //    app.Dispose();
            //}

            var books = Excel.Workbook.GetAllOpen();

            var sb = new StringBuilder();

            foreach (var workbook in books)
            {
                sb.AppendLine(workbook.FullName);

                workbook.Dispose();
            }

            MessageBox.Show("Currently open workbooks:\r\n\r\n" + sb.ToString(), "Excel Info");
        }

        private void CreateEmail(STC.Automation.Office.Outlook.Application outlook)
        {
            using (var msg = (STC.Automation.Office.Outlook.MailItem)outlook.CreateItem(STC.Automation.Office.Outlook.Enums.ItemType.MailItem))
            {
                using (var recipient = msg.Recipients.Add("to@example.com"))
                    recipient.Type = (long)STC.Automation.Office.Outlook.Enums.MailRecipientType.To;
                using (var recipient = msg.Recipients.Add("cc@example.com"))
                    recipient.Type = (long)STC.Automation.Office.Outlook.Enums.MailRecipientType.CC;
                msg.Subject = "This is the subject";
                msg.Body = "This is the body of the email.";
                msg.Importance = STC.Automation.Office.Outlook.Enums.Importance.High;
                if (File.Exists(@"C:\test.txt"))
                    using (var attachment = msg.Attachments.Add(@"C:\test.txt")) ;
                msg.Recipients.ResolveAll();
                msg.Display();
                msg.Save();
                //msg.Send();
            }
        }

        private void btnNewOutlook_Click(object sender, EventArgs e)
        {
            using (var outlook = new STC.Automation.Office.Outlook.Application())
            {
                MessageBox.Show(outlook.Version.ToString());

                using (var myNameSpace = outlook.GetNameSpace("MAPI"))
                {
                    using (var folder = myNameSpace.GetDefaultFolder(STC.Automation.Office.Outlook.Enums.DefaultFolders.Outbox))
                        folder.Display();
                }

                CreateEmail(outlook);
                //outlook.Quit();
            }

        }

        private void btnOutlookExisting_Click(object sender, EventArgs e)
        {
            var apps = STC.Automation.Office.Outlook.Application.GetRunningApplications();
            MessageBox.Show("Existing Outlook instances: " + apps.Count.ToString() + "\n" + (apps.Count == 0 ? "A new window will be created." : ""));
            if (apps.Count == 0)
                apps.Add(new STC.Automation.Office.Outlook.Application());

            if (apps.Count > 0)
            {
                foreach (var app in apps)
                    app.Dispose();

                using (var outlook = STC.Automation.Office.Outlook.Application.GetOrCreateApplication())
                {
                    MessageBox.Show(outlook.Version.ToString());

                    using (var explorer = outlook.ActiveExplorer)
                    {
                        if (explorer != null)
                        {
                            explorer.Activate();
                            //using (var cur = explorer.CurrentFolder)
                            //    cur.Display();
                            using (var myNameSpace = outlook.GetNameSpace("MAPI"))
                            {
                                using (var folder = myNameSpace.GetDefaultFolder(STC.Automation.Office.Outlook.Enums.DefaultFolders.Outbox))
                                    explorer.CurrentFolder = folder;
                            }
                        }
                        else
                        {
                            using (var myNameSpace = outlook.GetNameSpace("MAPI"))
                            {
                                using (var folder = myNameSpace.GetDefaultFolder(STC.Automation.Office.Outlook.Enums.DefaultFolders.Outbox))
                                    folder.Display();
                            }
                        }
                    }

                    CreateEmail(outlook);
                    //outlook.Quit();
                }
            }
        }

        private void btnOutlookProcess_Click(object sender, EventArgs e)
        {
            using (var outlook = STC.Automation.Office.Outlook.Application.FromProcess(System.Diagnostics.Process.GetProcessesByName("outlook")[0]))
            {
                if (outlook != null)
                {
                    using (var explorer = outlook.ActiveExplorer)
                    {
                        if (explorer != null)
                            explorer.Activate();
                    }
                }
                else
                    MessageBox.Show("Could not attach to existing outlook process");
            }
        }

        private void btnCreateOutlookMail_Click(object sender, EventArgs e)
        {
            using (var outlook = new STC.Automation.Office.Outlook.Application())
            {
                using (var mail = (STC.Automation.Office.Outlook.MailItem)outlook.CreateItem(STC.Automation.Office.Outlook.Enums.ItemType.MailItem))
                {
                    using (var recipient = mail.Recipients.Add("software@pittsh.com.au"))
                    {
                        recipient.Type = (long)STC.Automation.Office.Outlook.Enums.MailRecipientType.To;
                    }

                    mail.Subject = "Test email";
                    mail.Body = "Hello from automation!";

                    mail.Closing += Mail_Closing;
                    mail.Sending += Mail_Sending;

                    // Doing this has the same (or at least similar) effect as calling mail.Display(), in that it shows the message and prevents
                    // mail.Display(true) from making the inspector modal. This in turn means that the method used here for fielding events will
                    // not work.
                    // mail.Inspector.Activate();


                    // create temporary files for attachments as the method only accepts a filepath

                    // first attachment is just added and the attachment object disposed
                    var root = Path.Combine(Path.GetTempPath(), "STC.Automation.Office");
                    try { Directory.CreateDirectory(root); } catch { }
                    var filename = Path.Combine(root, "attached-image.jpg");
                    using (var writer = new FileStream(filename, FileMode.Create))
                        Properties.Resources.attached_image.Save(writer, System.Drawing.Imaging.ImageFormat.Jpeg);
                    mail.Attachments.Add(filename).Dispose();

                    // second attachment will be added inline in the message body
                    filename = Path.Combine(root, "embedded-image.png");
                    using (var writer = new FileStream(filename, FileMode.Create))
                        Properties.Resources.embedded_image.Save(writer, System.Drawing.Imaging.ImageFormat.Png);
                    using (var attachment = mail.Attachments.Add(filename))
                    {
                        // derived from https://stackoverflow.com/a/14052552
                        attachment.PropertyAccessor.SetProperty(Attachment.PR_ATTACH_MIME_TAG, "image/png");
                        attachment.PropertyAccessor.SetProperty(Attachment.PR_ATTACH_CONTENT_ID, "myContentId");
                    }
                    mail.HtmlBody = "<p>Hello from automation!</p><p><img src=cid:myContentId width=400 height=300></p>";

                    mail.Display(true);
                }
            }
        }

        private void Mail_Sending(object sender, ref bool cancel)
        {
            MessageBox.Show(this, "Mail sending");
        }

        private void Mail_Closing(object sender, ref bool cancel)
        {
            MessageBox.Show(this, "Mail closing");
        }
    }
}
