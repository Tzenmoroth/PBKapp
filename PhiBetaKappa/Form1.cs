using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;

namespace PhiBettaKappa
{
    public partial class Form1 : Form
    {
        private static float GPAmin = 3.65f;
        private static int creditsMin = 90;
        private static int onePageThreshold = 43; // MORE courses than this causes the student's transcript to be two pages on the document

        private static System.Drawing.Color FAIL_COLOR = System.Drawing.Color.Red;
        private static String[] nameLocator = { "Record", "of" };
        private static String[] termLocator = { "TERM" };
        private static String[] cumLocator = { "CUM" };
        private static String[] collegeLocator = {"College", ":"};
        private static String idLocatorStr = "Student No";
        private static String[] idLocator = {"Student", "No"};
        private static String[] progressLocator = {"In", "Progress", "Credits"};
        private static String[] overallLocator = {"OVERALL"};
        private static String[] transferLocator = {"TRANSFER", "CREDIT"};
        private static String[] institutionLocator = { "INSTITUTION", "CREDIT" };
        private static String[] transferTotalLocator = { "TOTAL", "TRANSFER" };
        private static String[] institutionTotalLocator = { "TOTAL", "INSTITUTION" };
        public static String[] subjects = { "TRSF", "GNRL", "GRAD", "OSSP", "VOSP", "BUCK", "HON", "HCOL", "AS", "IDM", "MATH", "STAT", "CS", "PHIL", "HUMN", "IHUM", 
                                         "REL", "HIST", "HST", "HP", "CLAS", "ARTH", "MU", "MUS", "DNCE", "ART", "ARTS", "IDFA", "FTSFilm", "FILM", "CT", "THE", "ENG", 
                                         "ENGL", "ENGS", "GLIT", "WLIT", "ESOL", "LANG", "HEBR", "ARBC", "CHIN", "JAPN", "RUSS", "SERB", "GERM", "GRK", "GKLT", "LAT", 
                                         "FREN", "ITAL", "SPAN", "PORT", "ASL", "CMSI", "AIS", "IS", "WP", "GRS", "HS", "ALAN", "CRES", "VS", "GSWS", "WGST", "WST", 
                                         "POLS", "PSCI", "IDSS", "ANTH", "EC", "ECON", "SOC", "LING", "CS&D", "CSD", "PSYC", "PSYS", "GEOG", "ISCI", "ASTR", "GEOL", 
                                         "PHYS", "PBIO", "BOT", "CHEM", "BIOL", "BCOR", "MMG", "ZOOL", "ENSC", "PSS", "BIOC", "BSCI", "CLBI", "CMB", "MCBI", "MICR", 
                                         "MPBP", "PSLB", "BIOE", "HLX", "E&ES", "NSCI", "AH", "ANAT", "ANES", "ANNB", "ANPS", "AT", "BIOS", "BMED", "BMT", "CTS", "DHYG", 
                                         "EMED", "EXMS", "FM", "FP", "GRMD", "GRNS", "GRNU", "HLTH", "HSCI", "MDMC", "MDPS", "MED", "MEDT", "MLRS", "MLS", "MVSR", "NEUR", 
                                         "NH", "NMT", "NURS", "OBGY", "ORTH", "PATH", "PED", "PH", "PHRM", "PRNU", "PT", "RAD", "RADT", "RMS", "RT", "SURG", "TENU", "CE", 
                                         "EE", "ME", "CEMS", "ENGR", "ES", "EMGT", "MATS", "CSYS", "ASTU", "TRC", "ASCI", "A&DS", "A&DH", "PS.", "ANPA", "ENVS", "FOR", 
                                         "HS.", "HUMS", "NR", "WFB", "WLB", "WR", "CALS", "AGRI", "AGHO", "AGBI", "AGE", "AGEC", "AGED", "AREC", "ANFS", "NFS", "NUSC", 
                                         "HN&F", "FS", "HEC", "RSEC", "H&RE", "CEC", "CDAE", "PRT", "RM", "MCSD", "TMCS", "CT&D", "COM", "MD", "SPCH", "BSAD", "PA", "MBA", 
                                         "EDFC", "HDFS", "SOSE", "SWSS", "ARED", "BSED", "ECHD", "ECSP", "EDAP", "EDAR", "EDCI", "EDCO", "EDEC", "EDEL", "EDFS", "EDHE", 
                                         "EDHI", "EDLI", "EDLP", "EDLS", "EDLT", "EDML", "EDMU", "EDOH", "EDPE", "EDRC", "EDRT", "EDSC", "EDSP", "EDSS", "EDTE", "EDUC", 
                                         "HEED", "MAED", "SPED", "FL.", "PE"/* W */, "PEP", "PEAC", "MS", "MSTD", "ZAST", "PSTG", "TECH", "VOTC", "EXP" };
        private static String[] grades = { "A+", "A", "A-", "B+", "B", "B-", "C+", "C", "C-", "D+", "D", "D-", "X", "XF", "AU",
                                         "INC", "P", "NP", "S", "U", "SP", "UP", "M", "W", "TR", "XC", "WP", "WF" };
        private static String[] semesters = { "Fall", "Spring", "Summer" };

        private static char[] removeable = { '*', '_', ':' };
        private static String digits = "0123456789";
        private static String courseDigits = "0123456789XL";

        // The transcripts, line by line
        private static List<List<Token>> transcriptLines;
        // The ID numbers
        private static List<String> IDnums;
        // The lines of transcripts on which a new transcript begins
        private static List<int> startIndexes;
        // The students
        private static List<Student> students;

        public Form1()
        {
            InitializeComponent();
        }
        
        public Cell makeCell(String str, Boolean isItANumber = false)
        {
            return new Cell() { CellValue = new CellValue(str), DataType = new EnumValue<CellValues>(isItANumber ? CellValues.Number : CellValues.String) };
        }

        private void export_Click(object sender, EventArgs e)
        {
            if (students == null)
            {
                MessageBox.Show("No data to export yet!", "Error");
                return;
            }
            saveFileDialog1.Title = "Export Transcript Data";
            saveFileDialog1.Filter = "Excel spreadsheets (*.xlsx)|*.xlsx";
            saveFileDialog1.ShowDialog();
        }
            
        // Save the spreadsheet and document - Edit this to edit those files
        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            try
            {
                // This section creates the spreadsheet
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(saveFileDialog1.FileName, SpreadsheetDocumentType.Workbook))
                {
                    // create the workbook
                    spreadSheet.AddWorkbookPart();
                    spreadSheet.WorkbookPart.Workbook = new Workbook();     // create the worksheet
                    spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
                    spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet = new Worksheet();

                    // create sheet data
                    spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet.AppendChild(new SheetData());

                    // Header row
                    Row headerRow = new Row(new OpenXmlElement[]{
                        makeCell("College"),
                        makeCell("Print Order"),
                        makeCell("UVM ID"),
                        makeCell("Last Name"),
                        makeCell("Prefered Name"),
                        makeCell("GPA"),
                        makeCell("GPA Strength"),
                        makeCell(" "),
                        makeCell("Totals Trans"),
                        makeCell("UVM Finished"),
                        makeCell("GPA Credits"),
                        makeCell("IP Credits"),
                        makeCell("PBK Credits"),
                        makeCell("Grad Credits"),
                        makeCell(" "),
                        makeCell("LANG"),
                        makeCell("ARBC"),
                        makeCell("ASL"),
                        makeCell("CHIN"),
                        makeCell("FREN"),
                        makeCell("GERM"),
                        makeCell("GRK"),
                        makeCell("HEBR"),
                        makeCell("ITAL"),
                        makeCell("JAPN"),
                        makeCell("LAT"),
                        makeCell("PORT"),
                        makeCell("RUSS"),
                        makeCell("SPAN")
                    });

                    spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet.First().AppendChild(headerRow);

                    Row temp;
                    for (int i = 0; i < students.Count; i++)
                    {
                        temp = new Row();
                        temp.Append(makeCell(students[i].college[0]));
                        temp.Append(makeCell(i.ToString(), true)); // Print order
                        temp.Append(makeCell(students[i].ID, true));
                        temp.Append(makeCell(students[i].lastName));
                        temp.Append(makeCell(students[i].prefName));
                        temp.Append(makeCell(students[i].GPA.ToString(), true));
                        temp.Append(makeCell((100.0f * students[i].institution / students[i].credits).ToString(), true));
                        temp.Append(makeCell(""));
                        temp.Append(makeCell(students[i].transfer.ToString(), true)); // transfer
                        temp.Append(makeCell(students[i].institution.ToString(), true)); // institution
                        temp.Append(makeCell((students[i].institution + students[i].transfer).ToString(), true)); // inst + trans
                        temp.Append(makeCell(students[i].progress.ToString(), true)); // progress
                        temp.Append(makeCell((students[i].institution + students[i].progress).ToString(), true)); // inst + prog
                        temp.Append(makeCell(students[i].credits.ToString(), true)); // all three
                        temp.Append(makeCell(""));
                        temp.Append(makeCell(students[i].getNumbers("LANG"), false));
                        temp.Append(makeCell(students[i].getNumbers("ARBC"), false));
                        temp.Append(makeCell(students[i].getNumbers("ASL"), false));
                        temp.Append(makeCell(students[i].getNumbers("CHIN"), false));
                        temp.Append(makeCell(students[i].getNumbers("FREN"), false));
                        temp.Append(makeCell(students[i].getNumbers("GERM"), false));
                        temp.Append(makeCell(students[i].getNumbers("GRK"), false));
                        temp.Append(makeCell(students[i].getNumbers("HEBR"), false));
                        temp.Append(makeCell(students[i].getNumbers("ITAL"), false));
                        temp.Append(makeCell(students[i].getNumbers("JAPN"), false));
                        temp.Append(makeCell(students[i].getNumbers("LAT"), false));
                        temp.Append(makeCell(students[i].getNumbers("PORT"), false));
                        temp.Append(makeCell(students[i].getNumbers("RUSS"), false));
                        temp.Append(makeCell(students[i].getNumbers("SPAN"), false));
                        spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet.First().AppendChild(temp);
                    }

                    SetColumnWidth(spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet, 1, 16.0 * 1.09);
                    SetColumnWidth(spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet, 2, 10.2 * 1.09);
                    SetColumnWidth(spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet, 3, 10.2 * 1.09);
                    SetColumnWidth(spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet, 4, 10.2 * 1.09);
                    SetColumnWidth(spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet, 5, 14.0 * 1.09);
                    SetColumnWidth(spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet, 6, 5.0 * 1.09);
                    SetColumnWidth(spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet, 7, 12.0 * 1.09);
                    SetColumnWidth(spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet, 8, 8.4 * 1.09);
                    SetColumnWidth(spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet, 9, 10.71 * 1.09);
                    SetColumnWidth(spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet, 10, 12.71 * 1.09);
                    SetColumnWidth(spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet, 11, 10.71 * 1.09);
                    SetColumnWidth(spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet, 12, 8.71 * 1.09);
                    SetColumnWidth(spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet, 13, 10.43 * 1.09);
                    SetColumnWidth(spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet, 14, 11.14 * 1.09);
                    SetColumnWidth(spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet, 15, 8.4 * 1.09);

                    // save worksheet
                    spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet.Save();

                    // create the worksheet to workbook relation
                    spreadSheet.WorkbookPart.Workbook.AppendChild(new Sheets());
                    spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().AppendChild(new Sheet()
                    {
                        Id = spreadSheet.WorkbookPart.GetIdOfPart(spreadSheet.WorkbookPart.WorksheetParts.First()),
                        SheetId = 1,
                        Name = "Sheet 1"
                    });

                    spreadSheet.WorkbookPart.Workbook.Save();
                }
                // This section creates the document
                using (WordprocessingDocument package = WordprocessingDocument.Create(saveFileDialog1.FileName.Replace(".xlsx", ".docx"), WordprocessingDocumentType.Document))
                {
                    // Add a new main document part. 
                    package.AddMainDocumentPart();

                    // Create the Document DOM. 
                    package.MainDocumentPart.Document = new Document();

                    // The indexes of the Body elements in the document
                    List<int> bodyIndexes = new List<int>();
                    int position = 0;

                    for (int i = 0; i < students.Count; i++)
                    {
                        package.MainDocumentPart.Document.Append(
                            new Body(
                                new SectionProperties(new PageSize() { Width = (UInt32Value)15840U, Height = (UInt32Value)12240U, Orient = PageOrientationValues.Landscape }, new DocumentFormat.OpenXml.Wordprocessing.Columns() { ColumnCount = 2 },
                                new PageMargin() { Top = 400, Right = Convert.ToUInt32(0.62 * 1440.0), Bottom = 350, Left = Convert.ToUInt32(0.62 * 1440.0), Header = (UInt32Value)450U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U }),
                                makePara(students[i].lastName + ", " + students[i].prefName, false, "32"),
                                makePara(students[i].ID),
                                makePara(students[i].GPA + " based on " + (100.0f * students[i].institution / students[i].credits) + " percent of credits toward graduation.", false, "24"),
                                makePara(" "),
                                makePara("PROGRESS TOWARDS GRADUATION", true),
                                makePara(" "),
                                makePara("CHECK DISTRIBUTION", true),
                                makePara(" "),
                                makePara("MATHEMATICS", true),
                                makePara(" "),
                                makePara("MATH OR HUMANITIES", true),
                                makePara(" "),
                                makePara("HUMANITIES", true),
                                makePara(" "),
                                makePara("HUMANITIES OR LITERATURE", true),
                                makePara(" "),
                                makePara("HUMANITIES OR FINE ARTS", true),
                                makePara(" "),
                                makePara("FINE ARTS", true),
                                makePara(" "),
                                makePara("LITERATURE OR WRITING", true),
                                makePara(" "),
                                makePara("LITERATURE", true),
                                makePara(" "),
                                makePara("LANGUAGES AND LITERATURE", true),
                                makePara(" "),
                                makePara("CHECK THIS ONE", true),
                                makePara(" "),
                                makePara("HUMANITIES OR SOCIAL SCIENCES", true),
                                makePara(" "),
                                makePara("SOCIAL SCIENCES", true),
                                makePara(" "),
                                makePara("SOCIAL SCIENCES OR NATURAL SCIENCES", true),
                                makePara(" "),
                                makePara("NATURAL SCIENCES", true),
                                makePara(" "),
                                makePara("NATURAL SCIENCES OR NON-CAS SCIENCE", true),
                                makePara(" "),
                                makePara("NON-CAS SCIENCE", true),
                                makePara(" "),
                                makePara("MEDICINE AND NURSING", true),
                                makePara(" "),
                                makePara("ENGINEERING", true),
                                makePara(" "),
                                makePara("CALS AND RSENR", true),
                                makePara(" "),
                                makePara("BUSINESS", true),
                                makePara(" "),
                                makePara("EDUCATION", true),
                                makePara(" "),
                                makePara("MISCELLANEOUS", true)
                        ));

                        // Record the index of this Body
                        bodyIndexes.Add(position);

                        package.MainDocumentPart.Document.Append(new DocumentFormat.OpenXml.Wordprocessing.Break() { Type = BreakValues.Page });
                        if(students[i].courseNums.Count <= onePageThreshold){
                            // This student's transcript is only one page, add another page break to make it two pages
                            package.MainDocumentPart.Document.Append(new DocumentFormat.OpenXml.Wordprocessing.Break() { Type = BreakValues.Page });
                            position += 3;
                        }
                        else
                        {
                            position += 2;
                        }
                    }

                    for (int j = 0; j < students.Count; j++) // For each STUDENT
                    {
                        for (int k = 0; k < 24; k++)
                        { // For each SUBJECT
                            foreach (String strI in students[j].coursesToString(23 - k))
                            { // For each COURSE in SUBJECT for STUDENT
                                package.MainDocumentPart.Document.ChildElements.GetItem(bodyIndexes[j]).InsertAt(makePara(strI), 52 - (2 * k)); // Print it
                            }
                        }
                    }

                    // Save changes to the main document part. 
                    package.MainDocumentPart.Document.Save();
                }
            } catch (IOException error) {
                MessageBox.Show("Could not save files. Make sure the files being saved are not already open and that you have permission to save in this location.", "Oops!");
                Debug.Print(error.StackTrace);
            }
        }

        private void SetColumnWidth(Worksheet worksheet, uint Index, double dwidth)
        {
            DocumentFormat.OpenXml.Spreadsheet.Columns cs = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Columns>();
            if (cs != null)
            {
                IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Column> ic = cs.Elements<DocumentFormat.OpenXml.Spreadsheet.Column>().Where(r => r.Min == Index).Where(r => r.Max == Index);
                if (ic.Count() > 0)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Column c = ic.First();
                    c.Width = dwidth;
                }
                else
                {
                    DocumentFormat.OpenXml.Spreadsheet.Column c = new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = Index, Max = Index, Width = dwidth, CustomWidth = true };
                    cs.Append(c);
                }
            }
            else
            {
                cs = new DocumentFormat.OpenXml.Spreadsheet.Columns();
                DocumentFormat.OpenXml.Spreadsheet.Column c = new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = Index, Max = Index, Width = dwidth, CustomWidth = true };
                cs.Append(c);
                worksheet.InsertAfter(cs, worksheet.GetFirstChild<SheetFormatProperties>());
            }
        }

        private Paragraph makePara(String text, Boolean bold = false, String size = "20", String font = "Times New Roman")
        {
            DocumentFormat.OpenXml.Wordprocessing.RunProperties runProp = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
            RunFonts rFont = new RunFonts();
            DocumentFormat.OpenXml.Wordprocessing.FontSize rSize = new DocumentFormat.OpenXml.Wordprocessing.FontSize();
            DocumentFormat.OpenXml.Wordprocessing.Bold rBold = new DocumentFormat.OpenXml.Wordprocessing.Bold();
            rBold.Val = OnOffValue.FromBoolean(true);
            rFont.Ascii = font;
            rSize.Val = size;
            runProp.Append(rFont);
            runProp.Append(rSize);
            if (bold) runProp.Append(rBold);

            SpacingBetweenLines spacing = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto, Before = "0", After = "0" };
            ParagraphProperties paraProps = new ParagraphProperties();
            paraProps.Append(spacing);

            return new Paragraph(paraProps, new DocumentFormat.OpenXml.Wordprocessing.Run(runProp, new DocumentFormat.OpenXml.Wordprocessing.Text() { Text = text, Space = SpaceProcessingModeValues.Preserve }));
        }

        private void load_Click(object sender, EventArgs e)
        {
            // Open the document
            openFileDialog1.Title = "Open Document";
            openFileDialog1.Filter = "Word documents (*.docx)|*.docx|All files (*.*)|*.*";
            openFileDialog1.ShowDialog();
        }

        // Load the file
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            String inputText = TextFromWord(openFileDialog1.FileName);
            if (inputText == null)
            {
                MessageBox.Show("Failed to open the document. Make sure the document is not open.", "Error");
                return;
            }
            if (inputText.Length == 0)
            {
                MessageBox.Show("Failed to open the document. Make sure the document is not open.", "Error");
                return;
            }
            // Reset variables
            transcriptLines = new List<List<Token>>();
            IDnums = new List<String>();
            startIndexes = new List<int>();
            students = new List<Student>();
            treeView1.Nodes.Clear();
            // Build transcripts
            String[] transcriptRegex = Regex.Split(inputText, "\r\n|\r|\n");
            String newID;
            int temp;
            for (int i = 0; i < transcriptRegex.Length; i++)
            {
                if ((temp = transcriptRegex[i].LastIndexOf(idLocatorStr)) < 0) continue;
                newID = transcriptRegex[i].Substring(temp + idLocatorStr.Length);
                newID = newID.Substring(newID.IndexOf(' '));
                newID = newID.TrimStart();
                newID = newID.Substring(0, newID.IndexOf(' '));
                if (IDnums.Count > 0 && IDnums[IDnums.Count - 1].Equals(newID)) continue;
                IDnums.Add(newID);
                students.Add(new Student());
                students[students.Count - 1].ID = newID;
                startIndexes.Add(i);
                treeView1.Nodes.Add(newID);
            }
            parseTranscript(transcriptRegex);
            // Go through each transcript, building data
            buildData();
            // Check criteria for acceptance
            // checkReqs();
        }

        public static string TextFromWord(String filepath)
        {
            const string wordmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            StringBuilder textBuilder = new StringBuilder();
            try
            {
                using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(filepath, false))
                {
                    // Manage namespaces to perform XPath queries.  
                    NameTable nt = new NameTable();
                    XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                    nsManager.AddNamespace("w", wordmlNamespace);

                    // Get the document part from the package.  
                    // Load the XML in the document part into an XmlDocument instance.  
                    XmlDocument xdoc = new XmlDocument(nt);
                    xdoc.Load(wdDoc.MainDocumentPart.GetStream());

                    XmlNodeList paragraphNodes = xdoc.SelectNodes("//w:p", nsManager);
                    foreach (XmlNode paragraphNode in paragraphNodes)
                    {
                        XmlNodeList textNodes = paragraphNode.SelectNodes(".//w:t", nsManager);
                        foreach (System.Xml.XmlNode textNode in textNodes)
                        {
                            textBuilder.Append(textNode.InnerText);
                        }
                        textBuilder.Append(Environment.NewLine);
                    }
                }
            }catch(Exception error){
                Debug.Print(error.StackTrace);
                return null;
            }
            return textBuilder.ToString();
        }

        // Lines of document ---> Arrays of labelled tokens
        public void parseTranscript(String[] transcriptRegex)
        {
            int start, end, tempInt, tempInt2;
            Boolean peW = false;
            String tempStr;
            List<String> data;
            for (int i = 0; i < startIndexes.Count; i++) // For each transcript
            {
                start = startIndexes[i];
                end = i == startIndexes.Count - 1 ? transcriptRegex.Length : startIndexes[i + 1];
                for (int j = start; j < end; j++) // For each line
                {
                    while (j >= transcriptLines.Count) transcriptLines.Add(new List<Token>());
                    data = transcriptRegex[j].Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList<String>();
                    for (int k = 0; k < data.Count; k++) // For each word
                    {
                        if(isRemoveable(data[k])){
                            continue;
                        }
                        if(k < data.Count - idLocator.Length - 1 && data[k].Contains(idLocator[0]) && data[k + 1].Contains(idLocator[1]))
                        {
                            transcriptLines[j].Add(new Token(data[k] + " " + data[k + 1], TokenType.IDtag));
                            transcriptLines[j].Add(new Token(data[k + 2], TokenType.IDnum));
                            data.RemoveAt(k + 1);
                            data.RemoveAt(k + 2 - 1);
                            continue;
                        }
                        if (k < data.Count - progressLocator.Length && data[k].Contains(progressLocator[0]) && data[k + 1].Contains(progressLocator[1]) && data[k + 2].Contains(progressLocator[2]))
                        {
                            transcriptLines[j].Add(new Token(data[k] + " " + data[k + 1] + " " + data[k + 2], TokenType.ProgressTag));
                            data.RemoveAt(k + 1);
                            data.RemoveAt(k + 2 - 1);
                            continue;
                        }
                        if (k < data.Count - overallLocator.Length && data[k].Contains(overallLocator[0]))
                        {
                            transcriptLines[j].Add(new Token(data[k], TokenType.OverallTag));
                            continue;
                        }
                        if (k < data.Count - termLocator.Length && data[k].Contains(termLocator[0]))
                        {
                            transcriptLines[j].Add(new Token(data[k], TokenType.TermTag));
                            continue;
                        }
                        if (k < data.Count - cumLocator.Length && data[k].Contains(cumLocator[0]))
                        {
                            transcriptLines[j].Add(new Token(data[k], TokenType.CumTag));
                            continue;
                        }
                        if (k < data.Count - collegeLocator.Length && data[k].Contains(collegeLocator[0]) && data[k + 1].Contains(collegeLocator[1]))
                        {
                            transcriptLines[j].Add(new Token(data[k] + " " + data[k + 1], TokenType.CollegeTag));
                            continue;
                        }
                        if (k < data.Count - transferLocator.Length && data[k].Contains(transferLocator[0]) && data[k + 1].Contains(transferLocator[1]))
                        {
                            transcriptLines[j].Add(new Token(data[k] + " " + data[k + 1], TokenType.TransferTag));
                            data.RemoveAt(k + 1);
                            continue;
                        }
                        if (k < data.Count - nameLocator.Length && data[k].Contains(nameLocator[0]) && data[k + 1].Contains(nameLocator[1]))
                        {
                            transcriptLines[j].Add(new Token(data[k] + " " + data[k + 1], TokenType.NameTag));
                            data.RemoveAt(k + 1);
                            continue;
                        }
                        if (k < data.Count - institutionLocator.Length && data[k].Contains(institutionLocator[0]) && data[k + 1].Contains(institutionLocator[1]))
                        {
                            transcriptLines[j].Add(new Token(data[k] + " " + data[k + 1], TokenType.InstitutionTag));
                            data.RemoveAt(k + 1);
                            continue;
                        }
                        if (k < data.Count - institutionTotalLocator.Length && data[k].Contains(institutionTotalLocator[0]) && data[k + 1].Contains(institutionTotalLocator[1]))
                        {
                            transcriptLines[j].Add(new Token(data[k] + " " + data[k + 1], TokenType.InsTotalTag));
                            data.RemoveAt(k + 1);
                            continue;
                        }
                        if (k < data.Count - transferTotalLocator.Length && data[k].Contains(transferTotalLocator[0]) && data[k + 1].Contains(transferTotalLocator[1]))
                        {
                            transcriptLines[j].Add(new Token(data[k] + " " + data[k + 1], TokenType.TransTotalTag));
                            data.RemoveAt(k + 1);
                            continue;
                        }
                        if (k < data.Count - 2 && subjects.Contains<String>(data[k]))
                        {
                            peW = false;
                            if(!isCourseNumMaybe(data[k + 1])){
                                if(data[k].Equals("PE") && data[k + 1].Equals("W")) {
                                    peW = true;
                                }
                                else
                                {
                                    transcriptLines[j].Add(new Token(data[k], TokenType.Null));
                                    continue;
                                }

                            }
                            tempInt = tempInt2 = -1;
                            tempStr = "";
                            for (int tk = k; tk < data.Count; tk++) {
                                if (tempInt2 < 0 && isNumWithADot(data[tk])) {
                                    tempInt2 = tk;
                                    transcriptLines[j].Add(new Token(data[tk], TokenType.Cred));
                                    continue;
                                }
                                if (grades.Contains<String>(data[tk]) && isNumWithADot(data[tk - 1]))
                                {
                                    tempInt = tk;
                                    transcriptLines[j].Add(new Token(data[tk], TokenType.Grade));
                                    break;
                                }
                                if (tk < data.Count - 1 && data[tk].Equals("IN") && data[tk + 1].Equals("PROGRESS"))
                                {
                                    tempInt = tk;
                                    transcriptLines[j].Add(new Token("IN PROGRESS", TokenType.Grade));
                                    data.RemoveAt(tk + 1);
                                    break;
                                }
                            }
                            if (tempInt < k) { // No grade found for some reason
                                transcriptLines[j].Add(new Token(data[k], TokenType.CourseSub));
                                if(tempInt2 > 0) data.RemoveAt(tempInt2 - 1);
                                continue;
                            }
                            if (tempInt - k < 2 && tempInt2 > 0) { // Credits but no class number or name for some reason
                                transcriptLines[j].Add(new Token(data[k], TokenType.CourseSub));
                                data.RemoveAt(tempInt2);
                                data.RemoveAt(tempInt);
                                continue;
                            }
                            if (tempInt - k < 1 && tempInt2 < 0) { // Nothing between course name and grade for some reason
                                transcriptLines[j].Add(new Token(data[k], TokenType.CourseSub));
                                data.RemoveAt(tempInt);
                                continue;
                            }
                            if (peW)
                            {
                                transcriptLines[j].Add(new Token(data[k] + data[k + 1], TokenType.CourseSub));
                                transcriptLines[j].Add(new Token(data[k + 2], TokenType.CourseNum));
                                for (int tk2 = k + 3; tk2 < tempInt; tk2++) if (tk2 != tempInt2) tempStr += data[tk2] + " ";
                                tempStr = tempStr.TrimEnd();
                                transcriptLines[j].Add(new Token(tempStr, TokenType.CourseTitle));
                                continue;
                            }
                            transcriptLines[j].Add(new Token(data[k], TokenType.CourseSub));
                            transcriptLines[j].Add(new Token(data[k + 1], TokenType.CourseNum));
                            // BETWEEN course name and number AND grade, everythng that is NOT Credits is title
                            for (int tk2 = k + 2; tk2 < tempInt; tk2++) if(tk2 != tempInt2) tempStr += data[tk2] + " ";
                            tempStr = tempStr.TrimEnd();
                            transcriptLines[j].Add(new Token(tempStr, TokenType.CourseTitle));
                            continue;
                        }
                        if(isNumWithADot(data[k])){
                            transcriptLines[j].Add(new Token(data[k], TokenType.Value));
                            continue;
                        }
                        transcriptLines[j].Add(new Token(data[k], TokenType.Null));
                    }
                }
            }
        }

        // Arrays of labelled tokens ---> Student objects filled with data
        // Fill the treeView with data for each student
        public void buildData(){
            int start, end;
            int tempInt;
            String tempStr;
            List<int> indexes;
            List<String> tempData = new List<String>();
            List<String> tempNames = new List<String>();
            List<String> courseNames = new List<String>();
            List<String> courseNums = new List<String>();
            List<String> courseTitles = new List<String>();
            List<String> grades = new List<String>();
            List<String> creds = new List<String>();
            for (int i = 0; i < startIndexes.Count; i++) // For each transcript
            {
                courseNames.Clear();
                courseNums.Clear();
                courseTitles.Clear();
                grades.Clear();
                creds.Clear();
                start = startIndexes[i];
                end = (i == startIndexes.Count - 1 ? transcriptLines.Count : startIndexes[i + 1]);
                for (int j = start; j < end; j++) // For each line
                {
                    tempData.Clear();
                    if ((indexes = Token.indexesOfType(transcriptLines[j], TokenType.OverallTag)).Count > 0) {
                        for (int k = indexes[0] + 1; k < indexes[0] + 5 && k < transcriptLines[j].Count; k++)
                        {
                            if (transcriptLines[j][k].type == TokenType.Value) {
                                tempData.Add(transcriptLines[j][k].data);
                            } else {
                                tempData.Add("");
                            }
                        }
                        while (tempData.Count < 4) tempData.Add("");
                        students[i].earned = (int)Convert.ToDouble(tempData[0]);
                        students[i].hours = (int)Convert.ToDouble(tempData[1]);
                        students[i].GPA = (float)Convert.ToDouble(tempData[3]);
                        treeView1.Nodes[i].Nodes.Add("Earned Hrs: " + tempData[0] + ", GPA Hrs: " + tempData[1] + ", Points: " + tempData[2]);
                        treeView1.Nodes[i].Nodes.Add("GPA: " + tempData[3]);
                    }
                    if ((indexes = Token.indexesOfType(transcriptLines[j], TokenType.ProgressTag)).Count > 0)
                    {
                        tempData.Insert(0, "");
                        if (indexes[0] < transcriptLines[j].Count - 1)
                        {
                            if (transcriptLines[j][indexes[0] + 1].type == TokenType.Value)
                            {
                                tempData[0] = transcriptLines[j][indexes[0] + 1].data;
                            }
                        }
                        students[i].progress = (int)Convert.ToDouble(tempData[0]);
                        treeView1.Nodes[i].Nodes.Add("In Progress: " + tempData[0]);
                    }
                    if (students[i].prefName.Length == 0 && (indexes = Token.indexesOfType(transcriptLines[j], TokenType.NameTag)).Count > 0)
                    {
                        tempNames = new List<string>();
                        if (indexes[0] < transcriptLines[j].Count - 1 && transcriptLines[j][indexes[0] + 1].type == TokenType.Null)
                        {
                            tempNames.Add(transcriptLines[j][indexes[0] + 1].data);
                            for (int ki = indexes[0] + 2; ki < transcriptLines[j].Count; ki++)
                            {
                                if (transcriptLines[j][ki].data.Equals("U") && ki + 1 < transcriptLines[j].Count && transcriptLines[j][ki + 1].data.Equals("N"))
                                {
                                    break;
                                }
                                tempNames.Add(transcriptLines[j][ki].data);
                            }
                            students[i].prefName = tempNames[0];
                            for (int ki = 1; ki < tempNames.Count - 2; ki++) students[i].prefName = students[i].prefName + " " + tempNames[ki];
                            students[i].lastName = tempNames[tempNames.Count - 1];
                            treeView1.Nodes[i].Nodes.Add("Pref Name: " + students[i].prefName + ", Last Name: " + students[i].lastName);
                        }
                    }
                    if ((indexes = Token.indexesOfType(transcriptLines[j], TokenType.CollegeTag)).Count > 0)
                    {
                        tempInt = indexes[0]; // Will be the index of the end of college name
                        for(int ki = indexes[0] + 1; ki < transcriptLines[j].Count; ki++, tempInt++){
                            if(transcriptLines[j][ki].type != TokenType.Null) break;
                        }
                        if(tempInt > indexes[0]){
                            tempStr = transcriptLines[j][indexes[0] + 1].data;
                            for (int ji = indexes[0] + 2; ji <= tempInt; ji++)
                            {
                                tempStr += " " + transcriptLines[j][ji].data;
                            }
                            students[i].college.Add(Student.collegeName(tempStr));
                            treeView1.Nodes[i].Nodes.Add("College: " + students[i].college[students[i].college.Count - 1]);
                        }
                    }
                    if ((indexes = Token.indexesOfType(transcriptLines[j], TokenType.TransTotalTag)).Count > 0)
                    {
                        if (indexes[0] < transcriptLines[j].Count - 1 && transcriptLines[j][indexes[0] + 1].type == TokenType.Value)
                        {
                            students[i].transfer = (int)Convert.ToDouble(transcriptLines[j][indexes[0] + 1].data);
                            treeView1.Nodes[i].Nodes.Add("Transfer: " + transcriptLines[j][indexes[0] + 1].data);
                        }
                    }
                    if ((indexes = Token.indexesOfType(transcriptLines[j], TokenType.InsTotalTag)).Count > 0)
                    {
                        if (indexes[0] < transcriptLines[j].Count - 1 && transcriptLines[j][indexes[0] + 1].type == TokenType.Value)
                        {
                            students[i].institution = (int)Convert.ToDouble(transcriptLines[j][indexes[0] + 1].data);
                        }
                    }
                    if ((indexes = Token.indexesOfType(transcriptLines[j], TokenType.CourseSub)).Count > 0)
                    {
                        for (int ki = 0; ki < indexes.Count; ki++) courseNames.Add(transcriptLines[j][indexes[ki]].data);
                    }
                    if ((indexes = Token.indexesOfType(transcriptLines[j], TokenType.CourseNum)).Count > 0)
                    {
                        for (int ki = 0; ki < indexes.Count; ki++) courseNums.Add(transcriptLines[j][indexes[ki]].data);
                    }
                    if ((indexes = Token.indexesOfType(transcriptLines[j], TokenType.CourseTitle)).Count > 0)
                    {
                        for (int ki = 0; ki < indexes.Count; ki++) courseTitles.Add(transcriptLines[j][indexes[ki]].data);
                    }
                    if ((indexes = Token.indexesOfType(transcriptLines[j], TokenType.Grade)).Count > 0)
                    {
                        for (int ki = 0; ki < indexes.Count; ki++) grades.Add(transcriptLines[j][indexes[ki]].data);
                    }
                    if ((indexes = Token.indexesOfType(transcriptLines[j], TokenType.Cred)).Count > 0)
                    {
                        for (int ki = 0; ki < indexes.Count; ki++) creds.Add(transcriptLines[j][indexes[ki]].data);
                    }
                    while (creds.Count < courseNames.Count)
                    {
                        creds.Add("0");
                    }
                    while (courseTitles.Count < courseNames.Count)
                    {
                        courseTitles.Add("No Course Title");
                    }
                    while (grades.Count < courseNames.Count)
                    {
                        grades.Add("No Grade");
                    }
                    while (courseNums.Count < courseNames.Count)
                    {
                        courseNums.Add("???");
                    }
                }
                // End of transcript stuff
                students[i].credits = students[i].earned + students[i].progress;
                sortCourseNames(courseNames, courseTitles, courseNums, grades, creds);
                students[i].courseSubj = new List<String>(courseNames);
                students[i].courseNums = new List<String>(courseNums);
                students[i].grades = new List<String>(grades);
                students[i].titles = new List<String>(courseTitles);
                students[i].creds = new List<String>(creds);
                treeView1.Nodes[i].Nodes.Add("Courses");
                for (int ki = 0; ki < courseNames.Count; ki++) treeView1.Nodes[i].Nodes[treeView1.Nodes[i].Nodes.Count - 1].Nodes.Add(
                    courseNames[ki] + " " + courseNums[ki] + "   " + courseTitles[ki] + "   " + grades[ki]);
            }
        }

        // Check some requirements
        public void checkReqs()
        {
            Boolean d1, d2;
            for (int i = 0; i < students.Count; i++) // For each student
            {
                if (students[i].GPA < GPAmin)
                {
                    treeView1.Nodes[i].ForeColor = FAIL_COLOR;
                    treeView1.Nodes[i].Text += " (Insufficient GPA)";
                    continue;
                }
                if (students[i].credits < creditsMin)
                {
                    treeView1.Nodes[i].ForeColor = FAIL_COLOR;
                    treeView1.Nodes[i].Text += " (Insufficient Credits)";
                    continue;
                }
                d1 = d2 = false;
                for (int j = 0; j < students[i].titles.Count; j++)
                {
                    if (students[i].titles[j].Contains("D1")) d1 = true;
                    if (students[i].titles[j].Contains("D2")) d2 = true;
                }
                if (!d1)
                {
                    treeView1.Nodes[i].ForeColor = FAIL_COLOR;
                    treeView1.Nodes[i].Text += " (No D1 Course)";
                    continue;
                }
                if (!d2)
                {
                    treeView1.Nodes[i].ForeColor = FAIL_COLOR;
                    treeView1.Nodes[i].Text += " (No D2 Course)";
                    continue;
                }
            }
        }

        // Sort coursesToSort (subject and number) while making identical changes to titleList (titles)
        private static void sortCourseNames(List<String> coursesToSort, List<String> titleList, List<String> numList, List<String> gradeList, List<String> credList)
        {
            int n = coursesToSort.Count;
            String temp;
            Boolean swapped = true;
            while(swapped){
               swapped = false;
               for(int i = 1; i < n; i++){
                   if (compareCourseNames(coursesToSort[i], coursesToSort[i - 1]))
                   {
                      temp = coursesToSort[i - 1];
                      coursesToSort[i - 1] = coursesToSort[i];
                      coursesToSort[i] = temp;
                      temp = titleList[i - 1];
                      titleList[i - 1] = titleList[i];
                      titleList[i] = temp;
                      temp = gradeList[i - 1];
                      gradeList[i - 1] = gradeList[i];
                      gradeList[i] = temp;
                      temp = credList[i - 1];
                      credList[i - 1] = credList[i];
                      credList[i] = temp;
                      temp = numList[i - 1];
                      numList[i - 1] = numList[i];
                      numList[i] = temp;
                      swapped = true;
                  }
               }
               n = n - 1;
            }
        }

        // Return true if the Top name comes first
        private static Boolean compareCourseNames(String nameTop, String nameBottom)
        {
            String[] nameTopParsed = nameTop.Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            String[] nameBottomParsed = nameBottom.Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            int indexTop = -1, indexBottom = -1;
            for (int i = 0; i < subjects.Length; i++)
            {
                if (nameTopParsed[0].Equals(subjects[i])) indexTop = i;
                if (nameBottomParsed[0].Equals(subjects[i])) indexBottom = i;
            }
            if (indexTop > indexBottom) return false;
            if (indexTop < indexBottom) return true;
            if (nameTopParsed.Length < 2) return true;
            if (nameBottomParsed.Length < 2) return false;
            return courseNumAsInt(nameTopParsed[1]) < courseNumAsInt(nameBottomParsed[1]);
        }

        // Turn a course name into an integer
        private static int courseNumAsInt(String numStr)
        {
            int output = 0;
            for (int i = 0; i < numStr.Length; i++)
            {
                output += (int)Math.Pow(10, i) * Math.Max(digits.IndexOf(numStr[numStr.Length - 1 - i]), 0);
            }
            return output;
        }

        // Determine whether the string consists of digits and exactly one decimal point
        private static Boolean isNumWithADot(String strInput)
        {
            Boolean foundDot = false;
            for (int i = 0; i < strInput.Length; i++)
            {
                if (strInput[i] == '.')
                {
                    if (!foundDot)
                    {
                        foundDot = true;
                        continue;
                    }
                    return false;
                }
                if (digits.Contains(strInput[i])) continue;
                return false;
            }
            return foundDot;
        }

        private static Boolean isCourseNumMaybe(String strInput)
        {
            if (strInput.Length != 3) return false;
            for (int i = 0; i < strInput.Length; i++) if (!courseDigits.Contains(strInput[i])) return false;
            return true;
        }

        public static String extractNumber(String strInput, Boolean mayHaveDot)
        {
            String intStr = "";
            Boolean foundDot = !mayHaveDot;

            for (int i = 0; i < strInput.Length; i++)
            {
                if (digits.Contains(strInput[i]))
                {
                    intStr += strInput[i];
                } else if (strInput[i] == '.' && !foundDot) {
                    intStr += strInput[i];
                    foundDot = true;
                }
            }

            return intStr;
        }
        
        public static Boolean isRemoveable(String input) // is the entire string removeable
        {
            for (int i = 0; i < input.Length; i++) if(!removeable.Contains<char>(input[i])) return false;
            return true;
        }

    }
}
