using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace CSATExcelSorting {
    public partial class CSATExcelSorting : Form {

        private string FilePath;
        private string FileFolder;
        private string FileName;
        private BindingList<Site> LocalSitesDataGrid = new BindingList<Site>();
        private BindingList<Site> RemoteSitesDataGrid = new BindingList<Site>();
        private Excel.Application excelApp;
        private string Year = DateTime.Today.Year.ToString().Remove(0, 2);
        private string Month = getMonthName(DateTime.Today.Month);
        private int ProgressBarMax = 0;
        private int Progress = 0;
        const double CurrentVersion = 2.3;

        public CSATExcelSorting() {
            addLocalResolutionGroups();
            addRemoteResolutionGroups();
            excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            excelApp.Visible = true;

            InitializeComponent();
            dgvLocalSites.DataSource = LocalSitesDataGrid;
            dgvRemoteSites.DataSource = RemoteSitesDataGrid;
        }

        /// <summary>
        ///     just setting some intial registry keys to be able to determine if the application is up to date or not.
        ///     this may be very archaic, but it prevents extended usage without my authorization
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CSATFixer_Load(object sender, EventArgs e) {

            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Joseph Simon");

            
            string oneYear = DateTime.Today.Date.Add(new TimeSpan(365, 0, 0, 0)).ToShortDateString();

            if (key == null) {
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Joseph Simon");
                key.SetValue("ExpDate", oneYear);
                key.SetValue("Version", CurrentVersion);
            } else {
                //parses the info stored in the registry to verify if the version is still valid
                string date = (string)key.GetValue("ExpDate");
                double version = double.Parse((string)key.GetValue("Version"));
                if (DateTime.Today.Date.ToShortDateString() == date && version == CurrentVersion) {
                    Form frmOutOfDate = new Form();
                    RichTextBox rtbOutOfDate = new RichTextBox();
                    rtbOutOfDate.SelectionAlignment = HorizontalAlignment.Center;
                    rtbOutOfDate.ReadOnly = true;
                    rtbOutOfDate.Parent = frmOutOfDate;
                    rtbOutOfDate.Dock = DockStyle.Fill;
                    rtbOutOfDate.Text = "SORRY!\nThis version is out of date!";
                    rtbOutOfDate.Font = new Font(FontFamily.GenericSansSerif, 30);
                    frmOutOfDate.StartPosition = FormStartPosition.CenterParent;
                    frmOutOfDate.Controls.Add(rtbOutOfDate);
                    frmOutOfDate.ControlBox = true;
                    frmOutOfDate.FormClosed += new FormClosedEventHandler(btnClose_Click);
                    frmOutOfDate.ShowDialog();
                } else if (version < CurrentVersion) {
                    Microsoft.Win32.Registry.CurrentUser.DeleteSubKey("SOFTWARE\\Joseph Simon");
                    key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Joseph Simon");
                    key.SetValue("ExpDate", oneYear);
                    key.SetValue("Version", CurrentVersion);
                }
            }
            key.Close();
        }

        /// <summary>
        ///     Creating each of the Local Sites and adding them to the initial List of sites.
        /// </summary>
        private void addLocalResolutionGroups() {
            var thisOne = new Site("This");
            thisOne.Groups.Add(new ResolutionGroup("THIS.DISP", "Dispatch"));
            LocalSitesDataGrid.Add(thisOne);        //index 0

            var ThatOne = new Site("That");
            ThatOne.Groups.Add(new ResolutionGroup("THAT.DISP", "Dispatch"));
            ThatOne.Groups.Add(new ResolutionGroup("THAT.KIOSK", "Kiosk"));
            LocalSitesDataGrid.Add(ThatOne);        //index 1

            var TheOther = new Site("TheOther");
            TheOther.Groups.Add(new ResolutionGroup("THEOTHER.DISP", "Dispatch"));
            TheOther.Groups.Add(new ResolutionGroup("THEOTHER.SITE", "Site"));
            LocalSitesDataGrid.Add(TheOther);       //index 2
        }

        /// <summary>
        ///     Creating each of the Remote Sites and adding them to the initial List of sites.
        /// </summary>
        private void addRemoteResolutionGroups() {
            var Foo = new Site("Foo");
            Foo.Groups.Add(new ResolutionGroup("FOO.DISP", "Dispatch"));
            Foo.Groups.Add(new ResolutionGroup("FOO.KIOSK", "Kiosk"));
            Foo.Groups.Add(new ResolutionGroup("FOO.SITE", "Site"));
            RemoteSitesDataGrid.Add(Foo);           //index 0

            var Bar = new Site("Bar");
            Bar.Groups.Add(new ResolutionGroup("BAR.DISP", "Dispatch"));
            Bar.Groups.Add(new ResolutionGroup("BAR.SITE", "Site"));
            RemoteSitesDataGrid.Add(Bar);           //index 1

            var Foobar = new Site("Foobar");
            Foobar.Groups.Add(new ResolutionGroup("FOOBAR.DISP", "Dispatch"));
            Foobar.Groups.Add(new ResolutionGroup("FOOBAR.SITE", "Site"));
            RemoteSitesDataGrid.Add(Foobar);         //index 2
        }

        /// <summary>
        ///     Shows the dialog to open a file, and only allows for specific files to be opened for parsing
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOpenCSAT_Click(object sender, EventArgs e) {
            openFileDialog.ShowDialog();
            FilePath = openFileDialog.FileName;
            //verify that the file chosen has the correct file name generated by the raw data file
            if (FilePath.Contains("ePoll CSAT")) {
                FileFolder = Path.GetDirectoryName(FilePath);
                FileName = Path.GetFileName(FilePath);
                tbLoadingProgress.Text = "Reading the CSATs from the selected file.";
                bwLoadingCSATs.RunWorkerAsync();
                frmLoadingProgress.ShowDialog();
            } else {
                var frmFileNameError = new Form();
                var tbWrongFileName = new TextBox();
                tbWrongFileName.Parent = frmFileNameError;
                tbWrongFileName.TextAlign = HorizontalAlignment.Center;
                tbWrongFileName.Dock = DockStyle.Fill;
                tbWrongFileName.ReadOnly = true;
                tbWrongFileName.Enabled = false;
                tbWrongFileName.Multiline = true;
                tbWrongFileName.Text = "Please open an \"ePoll CSAT\" file.";
                tbWrongFileName.Font = new Font(FontFamily.GenericSansSerif, 16);
                frmFileNameError.StartPosition = FormStartPosition.CenterParent;
                frmFileNameError.Size = new Size(200, 100);
                frmFileNameError.Controls.Add(tbWrongFileName);
                frmFileNameError.ControlBox = true;
                frmFileNameError.ShowDialog();
            }
        }
        
        /// <summary>
        ///     updates the ProgressBar while the process is being completed.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bwLoadingCSATs_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            if (progressBar.Maximum != ProgressBarMax)
                progressBar.Maximum = ProgressBarMax;
            // The progress percentage is a property of e
            progressBar.Value = e.ProgressPercentage;
        }

        /// <summary>
        ///     clean up after the background worker is complete reading and sorting the CSATs.
        ///     also makes more of the UI available.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bwLoadingCSATs_Completed(object sender, RunWorkerCompletedEventArgs e) {
            frmLoadingProgress.Close();

            //enable the data grid views
            dgvLocalSites.Enabled = true;
            dgvRemoteSites.Enabled = true;
            cbSelectAllLocalSites.Enabled = true;
            cbSelectAllRemoteSites.Enabled = true;
            dgvLocalSites.ClearSelection();
            dgvRemoteSites.ClearSelection();

            //allow the user to click on the BuildSheets buttons after this has run
            btnBuildSelectedSheets.Enabled = true;
            btnBuildAllLocalSites.Enabled = true;
            btnBuildAllRemoteSites.Enabled = true;
            btnBuildAllSites.Enabled = true;

            //disable the button to open a new file
            btnOpenCSAT.Enabled = false;
            progressBar.Value = (Progress = 0);
        }

        /// <summary>
        ///     This is kind of a long one, mainly for the long string of site queue names and the long switch statement.
        ///     This sorts each of the parsed CSATs into their individual sites, based on the names of the Assignment Group
        ///     received from the parsing of the ePoll CSAT file chosen.
        /// </summary>
        private void bwLoadingCSATs_DoWork(object sender, DoWorkEventArgs e) {
            Button sentObj = sender as Button;
            Excel.Workbook masterCSAT = excelApp.Workbooks.Open(FilePath);
            Excel.Sheets excelSheets = masterCSAT.Worksheets;
            Excel.Worksheet currentSheet = (Excel.Worksheet)excelSheets.Item[1];
            int lastRow = getMaxRow(currentSheet);
            int lastCol = getMaxCol(currentSheet);
            
            //these arrays are merely for searching and sorting purposes, that way it doesn't bother with the switch statement
            // or bother with creating any CSAT, unless it is part of the required sites.
            var LocalSiteQueues = new List<string>();
            foreach (Site site in LocalSitesDataGrid) {
                foreach (ResolutionGroup group in site.Groups) {
                    LocalSiteQueues.Add(group.Name);
                }
            }
            var RemoteSiteQueues = new List<string>();
            foreach (Site site in RemoteSitesDataGrid) {
                foreach (ResolutionGroup group in site.Groups) {
                    RemoteSiteQueues.Add(group.Name);
                }
            }

            ProgressBarMax = lastRow;
            List<CSAT> sats = new List<CSAT>();
            CSAT prevCSAT = new CSAT();
            CSAT newCSAT;
            string csatDate = stringCheckNull(currentSheet.get_Range("D3").Cells.Value);
            string[] date = csatDate.Split(new char[] { ' ', '/' });
            Month = getMonthName(int.Parse(date[0]), true);
            Year = date[2].Remove(0,2);

            for (int index = 3; index <= lastRow; index++) {
                //this array holds all of the information from each line of the ePoll excel sheet
                Array csatValues = (Array)currentSheet.get_Range("A" + index.ToString(), "AI" + index.ToString()).Cells.Value;
                //I have to run the check null on each of these parsed cells, due to being brought in from an excel sheet and 
                // the system is not perfect that outputs this data, so a lot of blank cells can, and will, be found.
                string assignmentGroup = stringCheckNull(csatValues.GetValue(1, 5));
                if (LocalSiteQueues.Contains(assignmentGroup) || RemoteSiteQueues.Contains(assignmentGroup)) {
                    //build the CSAT from the array of data provided by the CSAT sheet
                    newCSAT = new CSAT() {
                        TicketID = stringCheckNull(csatValues.GetValue(1, 1)),
                        AssignmentGroup = assignmentGroup,
                        ResolutionSpecialist = stringCheckNull(csatValues.GetValue(1, 7)),
                        CustomerComment = stringCheckNull(csatValues.GetValue(1, 24)),
                        Originator = stringCheckNull(csatValues.GetValue(1, 25)) + ' ' + stringCheckNull(csatValues.GetValue(1, 26)),
                        Q1Score = intCheckNull(csatValues.GetValue(1, 28)),
                        Q2Score = intCheckNull(csatValues.GetValue(1, 29)),
                        Q3Score = intCheckNull(csatValues.GetValue(1, 30)),
                        Q4Score = intCheckNull(csatValues.GetValue(1, 31)),
                        Q5Score = intCheckNull(csatValues.GetValue(1, 32)),
                        Q6Score = intCheckNull(csatValues.GetValue(1, 33)),
                        Q7Score = intCheckNull(csatValues.GetValue(1, 34)),
                        Q8Score = intCheckNull(csatValues.GetValue(1, 35))
                    };
                    //this verifies that the newly created CSAT is not a copy of the previous one, as the system outputs duplicates
                    if (newCSAT != prevCSAT) {
                        //foreach (Site site in LocalSitesDataGrid.Zip(RemoteSitesDataGrid, new List<Site> => { l, r } )) {
                        foreach (Site site in LocalSitesDataGrid.Union(RemoteSitesDataGrid)) {
                            if (site.hasGroup(assignmentGroup)) {
                                site.addCSAT(assignmentGroup, newCSAT);
                                break;
                            }
                        }
                        prevCSAT = newCSAT;
                    }
                }
                bwLoadingCSATs.ReportProgress(++Progress);
            }

            masterCSAT.Close();
        }

        /// <summary>
        ///     check for a blank cell value and return a string, if that is expected
        /// </summary>
        /// <param name="value"></param>
        /// <returns> a string of the cell contents</returns>
        private string stringCheckNull(object value) {
            if (value == null) {
                return "";
            }
            return value.ToString();
        }

        /// <summary>
        ///     check for a blank cell value and return an integer, if that is expected
        /// </summary>
        /// <param name="value"></param>
        /// <returns> 0 if the cell is blank or the object doesn't parse, otherwise returns an int value</returns>
        private int intCheckNull(object value) {
            int parsedNum;
            if (value == null) {
                return 0;
            }
            if (int.TryParse(value.ToString(), out parsedNum)) {
                return parsedNum;
            } else {
                return 0;
            }
        }

        /// <summary>
        ///     Returns the last row number that has any information in any cell of an excel sheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns> the last row number with any data </returns>
        private int getMaxRow(Excel.Worksheet worksheet) {
            int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            return lastRow;
        }

        /// <summary>
        ///     returns the last column number that has any information in any cell of an excel sheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns> the last column number with any data </returns>
        private int getMaxCol(Excel.Worksheet worksheet) {
            int lastCol = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
            return lastCol;
        }

        /// <summary>
        ///     returns a full column name using the column number as a basis
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns></returns>
        private string ColumnNumToString(int columnNumber) {
            int dividend = columnNumber;
            string strColumnName = "";
            int modulo;
            while (dividend > 0) {
                modulo = (dividend - 1) % 26;
                strColumnName = Convert.ToChar(65 + modulo).ToString() + strColumnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return strColumnName;
        }

        /// <summary>
        ///     parses the name of the ePoll CSAT file to get the date of the CSATs, then 
        ///     sends each of the selected Site Lists to have the CSAT workbooks created.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBuildSelectedSheets_Click(object sender, EventArgs e) {
            bwBuildingWorkbook.RunWorkerAsync(false);
            tbLoadingProgress.Text = "Building remote site workbook";
            frmLoadingProgress.ShowDialog();

            bwBuildingWorkbook.RunWorkerAsync(true);
            tbLoadingProgress.Text = "Building local site workbook";
            frmLoadingProgress.ShowDialog();
        }

        /// <summary>
        ///     sends ALL of the local sites to have the workbook created
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBuildAllLocalSites_Click(object sender, EventArgs e) {
            dgvLocalSites.SelectAll();
            cbSelectAllLocalSites.Checked = true;

            bwBuildingWorkbook.RunWorkerAsync(true);
            tbLoadingProgress.Text = "Building local site workbook";
            frmLoadingProgress.ShowDialog();
        }

        /// <summary>
        ///     sends ALL remote sites to have the workbook created
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBuildAllRemoteSites_Click(object sender, EventArgs e) {
            dgvRemoteSites.SelectAll();
            cbSelectAllRemoteSites.Checked = true;

            bwBuildingWorkbook.RunWorkerAsync(false);
            tbLoadingProgress.Text = "Building remote site workbook";
            frmLoadingProgress.ShowDialog();
        }

        /// <summary>
        ///     sends all sites to have workbooks created
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBuildAllSheets_Click(object sender, EventArgs e) {
            dgvRemoteSites.SelectAll();
            cbSelectAllRemoteSites.Checked = true;

            bwBuildingWorkbook.RunWorkerAsync(false);
            tbLoadingProgress.Text = "Building remote site workbook";
            frmLoadingProgress.ShowDialog();

            dgvLocalSites.SelectAll();
            cbSelectAllLocalSites.Checked = true;

            bwBuildingWorkbook.RunWorkerAsync(true);
            tbLoadingProgress.Text = "Building local site workbook";
            frmLoadingProgress.ShowDialog();
        }

        /// <summary>
        ///     gets all of the data bound items from the selected rows of the Remote data grid view
        /// </summary>
        /// <returns></returns>
        private List<Site> getRemoteSitesfromDGV() {
            var sites = new List<Site>();
            foreach (DataGridViewRow row in dgvRemoteSites.SelectedRows) {
                sites.Add((Site)row.DataBoundItem);
            }
            return sites;
        }

        /// <summary>
        ///     gets all of the data bound items from the selected rows of the local data grid view
        /// </summary>
        /// <returns></returns>
        private List<Site> getLocalSitesfromDGV() {
            var sites = new List<Site>();
            foreach (DataGridViewRow row in dgvLocalSites.SelectedRows) {
                sites.Add((Site)row.DataBoundItem);
            }
            return sites;
        }

        /// <summary>
        ///     update the Progress Bar, making sure that the maximum is set to the proper max
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bwBuildingWorkbook_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            if (progressBar.Maximum != ProgressBarMax)
                progressBar.Maximum = ProgressBarMax;
            // The progress percentage is a property of e
            progressBar.Value = e.ProgressPercentage;
        }
        /// <summary>
        ///     perform the final touches and resets after the background worker is completed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bwBuildingWorkbook_Completed(object sender, RunWorkerCompletedEventArgs e) {
            frmLoadingProgress.Close();
            btnOpenCSAT.Enabled = false;
            btnOpenFolder.Enabled = true;
            progressBar.Value = (Progress = 0);
        }

        /// <summary>
        ///     this is the logic used by the buttons to create a workbook and send it 
        ///     to be built and filled
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">
        ///    this receives the required boolean argument, which is used to 
        ///    tell the method whether or not the workbook creation request is local
        ///     or remote, and changes variables accordingly.
        /// </param>
        private void bwBuildingWorkbook_DoWork(object sender, DoWorkEventArgs e) {
            bool local = (bool)e.Argument;

            //condense the date into a format to be used in the building of the sheets
            string dateOfCSATS = "1 " + Month + " " + Year;
            string locality;
            List<Site> sites;

            if (local) {
                locality = "Local";
                sites = getLocalSitesfromDGV();
            } else {
                locality = "Remote";
                sites = getRemoteSitesfromDGV();
            }
            
            if (sites.Count > 0) {
                Excel.Workbook csatWorkbook = excelApp.Workbooks.Add();
                BuildWorkbook(ref sites, ref csatWorkbook, dateOfCSATS, local);

                if (sites.Count == LocalSitesDataGrid.Count) {
                    FileName = "\\" + locality + " Sites Report " + Month + Year + ".xlsx";
                } else {
                    FileName = "\\" + locality + " Sites Report " + Month + Year + " - Partial.xlsx";
                }
                csatWorkbook.Close(true, FileFolder + FileName, Type.Missing);
            } else {
                Form frmSiteSelectionError = new Form();
                RichTextBox rtbNoneSelected = new RichTextBox();
                rtbNoneSelected.SelectionAlignment = HorizontalAlignment.Center;
                rtbNoneSelected.ReadOnly = true;
                rtbNoneSelected.Enabled = false;
                rtbNoneSelected.Parent = frmSiteSelectionError;
                rtbNoneSelected.Dock = DockStyle.Fill;
                rtbNoneSelected.Text = "Please select some " + locality + " sites to continue.";
                rtbNoneSelected.Font = new Font(FontFamily.GenericSansSerif, 16);
                frmSiteSelectionError.StartPosition = FormStartPosition.CenterParent;
                frmSiteSelectionError.Size = new Size(200, 125);
                frmSiteSelectionError.Controls.Add(rtbNoneSelected);
                frmSiteSelectionError.ControlBox = true;
                frmSiteSelectionError.ShowInTaskbar = false;
                frmSiteSelectionError.ShowDialog();
            }

            if (sites.Count > 0) {
                FilePath = FileFolder + FileName;
            }
        }

        /// <summary>
        ///     create, build, and format the workbook for the given Site List
        /// </summary>
        /// <param name="sites"> List of sites to build teh workbook with </param>
        /// <param name="workbook"> the workbook to fill with the sites and CSATS </param>
        /// <param name="date"> date of the CSAT collection </param>
        /// <param name="remote"> bool to tell whether this function will be working with a remote site list or a normal site list </param>
        private void BuildWorkbook(ref List<Site> sites, ref Excel.Workbook workbook, string date, bool local = true) {
            int tableStyle = 1;
            Excel.Worksheet currentSheet;
            bool empty = true;
            ProgressBarMax = sites.Count;
            foreach (Site site in sites) {
                foreach (ResolutionGroup group in site.Groups) {
                    if (!group.isEmpty()) {
                        currentSheet = workbook.Worksheets.Add();
                        currentSheet.Name = group.Name;
                        //build the initial, blank template
                        BuildSheet(ref currentSheet);
                        //add data to the template
                        AddData(ref currentSheet, group);

                        //this is to add the final touches to the data on the sheet, so the data rectifies properly.
                        int lastRow = getMaxRow(currentSheet);
                        string strLastRow = lastRow.ToString();
                        currentSheet.Cells[2, 3] = "=AVERAGE(Q2:Q" + strLastRow + ")";
                        currentSheet.Cells[2, 3].NumberFormat = "0.00";
                        currentSheet.Cells[3, 3] = "=IF(C4-C5 > 0,AVERAGEIF(R2:R" + strLastRow + ", \">0\"), 0)";
                        currentSheet.Cells[3, 3].NumberFormat = "0.00";
                        currentSheet.Cells[4, 3] = "=SUM(S2:S" + strLastRow + ")";
                        currentSheet.Cells[5, 3] = "=COUNTIF(T2: T" + strLastRow + ", \"=yes\")";
                        currentSheet.Cells[7, 3] = "=SUM(Q2:Q" + strLastRow + ")";
                        currentSheet.Cells[7, 3].NumberFormat = "0.00";
                        currentSheet.Cells[8, 3] = "=SUM(R2:R" + strLastRow + ")";
                        currentSheet.Cells[8, 3].NumberFormat = "0.00";
                        currentSheet.Cells[2, 29] = "=COUNT(F2:F" + strLastRow + ")";
                        //grab the CSATs that were added to turn them into a table
                        Excel.Range table = currentSheet.Range[currentSheet.Cells[1, 6], currentSheet.Cells[lastRow, 25]];
                        //this is to rotate between the 7 "light" table styles available in Excel
                        if (tableStyle > 7) {
                            tableStyle = 1;
                        }
                        currentSheet.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, 
                                                        table, 
                                                        Type.Missing, 
                                                        Excel.XlYesNoGuess.xlYes, 
                                                        Type.Missing, 
                                                        "TableStyleLight" + tableStyle.ToString()).Name = group.Name + " Table";
                        tableStyle++;
                        //add the final touches to the formatting of the sheet, to make it look nice, after making it into a table
                        FormatSheet(ref currentSheet);
                        empty = false;
                    }
                }
                bwBuildingWorkbook.ReportProgress(++Progress);
            }
            if (!empty) {
                //now we add a new worksheet to place the date in and hide, so it can be referenced in the summary page.
                currentSheet = workbook.Worksheets.Add();
                excelApp.Windows.Application.ActiveWindow.SelectedSheets.Visible = false;
                currentSheet.Name = "Date";
                currentSheet.Cells[1, 1] = date;

                //add the final touch to the entire workbook before heading back to save and close the file.
                AddSummaryPage(ref sites, ref workbook, local);

                //this one line is to remove the "Sheet1" automatically created worksheet, when starting a new workbook.
                //this had to wait until now, due to having to have at least 1 active/visible worksheet in the workbook.
                workbook.Worksheets[workbook.Worksheets.Count].Delete();
            }
        }

        /// <summary>
        ///     Build the template for each worksheet in the workbook
        /// </summary>
        /// <param name="worksheet"></param>
        private void BuildSheet(ref Excel.Worksheet worksheet) {
            worksheet.Cells[2, 2] = "Raw Average";
            worksheet.Cells[3, 2] = "Filtered Average";
            worksheet.Cells[4, 2] = "Total Responses";
            worksheet.Cells[5, 2] = "Filtered Responses";
            worksheet.Cells[7, 2] = "Raw Sum";
            worksheet.Cells[8, 2] = "Filtered Sum";

            worksheet.Cells[1, 6] = "Ticket ID";
            worksheet.Cells[1, 7] = "Resolution Specialist";
            worksheet.Cells[1, 8] = "Originator";
            worksheet.Cells[1, 9] = "Q1 Score";
            worksheet.Cells[1, 10] = "Q2 Score";
            worksheet.Cells[1, 11] = "Q3 Score";
            worksheet.Cells[1, 12] = "Q4 Score";
            worksheet.Cells[1, 13] = "Q5 Score";
            worksheet.Cells[1, 14] = "Q6 Score";
            worksheet.Cells[1, 15] = "Q7 Score";
            worksheet.Cells[1, 16] = "Q8 Score";
            worksheet.Cells[1, 17] = "Average Score";
            worksheet.Cells[1, 18] = "Filtered Average Score";
            worksheet.Cells[1, 19] = "Count";
            worksheet.Cells[1, 20] = "Filtered (Y/N)";
            worksheet.Cells[1, 21] = "Filter Cause";
            worksheet.Cells[1, 22] = "Technician Kudo";
            worksheet.Cells[1, 23] = "Resolution Group";
            worksheet.Cells[1, 24] = "Customer Comment";
            worksheet.Cells[1, 25] = "Internal Comment";
            worksheet.Cells[2, 30] = "Resolution Group";
        }

        /// <summary>
        ///     Add the data from each csat located in the ResolutionGroup's list
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="group"></param>
        private void AddData(ref Excel.Worksheet worksheet, ResolutionGroup group) {
            foreach (CSAT csat in group.CSATS) {
                int row = group.CSATS.IndexOf(csat) + 2;
                Excel.Range cell = worksheet.Cells[row, 6];
                cell.Formula = csat.TicketID;
                cell = worksheet.Cells[row, 7];
                cell.Formula = csat.ResolutionSpecialist;
                cell = worksheet.Cells[row, 8];
                cell.Formula = csat.Originator;
                cell = worksheet.Cells[row, 9];
                cell.Formula = csat.Q1Score;
                cell = worksheet.Cells[row, 10];
                cell.Formula = csat.Q2Score;
                cell = worksheet.Cells[row, 11];
                cell.Formula = csat.Q3Score;
                cell = worksheet.Cells[row, 12];
                cell.Formula = csat.Q4Score;
                cell = worksheet.Cells[row, 13];
                cell.Formula = csat.Q5Score;
                cell = worksheet.Cells[row, 14];
                cell.Formula = csat.Q6Score;
                cell = worksheet.Cells[row, 15];
                cell.Formula = csat.Q7Score;
                cell = worksheet.Cells[row, 16];
                cell.Formula = csat.Q8Score;

                cell = worksheet.Cells[row, 17];
                cell.Formula = "=AVERAGEIF(" + ColumnNumToString(9) + row.ToString() + 
                                ":" + ColumnNumToString(16) + row.ToString() + ", \">0\")";
                cell.NumberFormat = "0.00";

                cell = worksheet.Cells[row, 18];
                cell.Formula = "=IF(T" + row + " = \"Yes\", 0, AVERAGEIF(I" + row + ":P" + row + ", \">0\"))";
                cell.NumberFormat = "0.00";

                cell = worksheet.Cells[row, 19];
                cell.Formula = 1;

                cell = worksheet.Cells[row, 20];
                cell.Validation.Add(Excel.XlDVType.xlValidateList, 
                                    Excel.XlDVAlertStyle.xlValidAlertStop, 
                                    Excel.XlFormatConditionOperator.xlBetween, 
                                    "Yes, No");
                cell.Validation.IgnoreBlank = true;
                cell.Validation.InCellDropdown = true;

                cell = worksheet.Cells[row, 21];
                cell.Validation.Add(Excel.XlDVType.xlValidateList,
                                    Excel.XlDVAlertStyle.xlValidAlertStop,
                                    Excel.XlFormatConditionOperator.xlBetween,
                                    "IT Process, Deadline Expectation, Technician, Incomplete");
                cell.Validation.IgnoreBlank = true;
                cell.Validation.InCellDropdown = true;

                cell = worksheet.Cells[row, 22];
                cell.Validation.Add(Excel.XlDVType.xlValidateList,
                                    Excel.XlDVAlertStyle.xlValidAlertStop,
                                    Excel.XlFormatConditionOperator.xlBetween,
                                    "Yes");
                cell.Validation.IgnoreBlank = true;
                cell.Validation.InCellDropdown = true;

                cell = worksheet.Cells[row, 23];
                cell.Formula = csat.AssignmentGroup;
                cell = worksheet.Cells[row, 24];
                cell.Formula = csat.CustomerComment;
                cell = worksheet.Cells[row, 28];
                cell.Formula = "=IF(T" + row + "=\"Yes\",1,0)";
                if (!csat.AboveTarget(4.50)) {
                    cell = worksheet.Range[worksheet.Cells[row, 6], worksheet.Cells[row, 28]];
                    cell.Interior.Color = Excel.XlRgbColor.rgbYellow;
                }
            }
        }

        /// <summary>
        ///     add the correct formatting to the sheet, to fix any issues after changing the data into a table
        /// </summary>
        /// <param name="worksheet"></param>
        private void FormatSheet(ref Excel.Worksheet worksheet) {
            worksheet.Cells.WrapText = true;

            worksheet.Rows.RowHeight = 30;
            worksheet.Rows.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
            worksheet.Rows[1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            worksheet.Rows.Font.Size = 11;
            worksheet.Rows[1].RowHeight = 40;
            worksheet.Rows[1].Font.Size = 10;
            worksheet.Rows[1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            
            worksheet.Columns[2].ColumnWidth = 10;
            worksheet.Columns[6].ColumnWidth = 12;
            worksheet.Columns[7].ColumnWidth = 12;
            worksheet.Columns[8].ColumnWidth = 20;
            worksheet.Columns[17].ColumnWidth = 7;
            worksheet.Columns[18].ColumnWidth = 9;
            worksheet.Columns[19].ColumnWidth = 6;
            worksheet.Columns[20].ColumnWidth = 8.5;
            worksheet.Columns[21].ColumnWidth = 10;
            worksheet.Columns[22].ColumnWidth = 10;
            worksheet.Columns[23].ColumnWidth = 12.5;
            worksheet.Columns[24].ColumnWidth = 64;
            worksheet.Columns[25].ColumnWidth = 64;
            worksheet.Columns[30].ColumnWidth = 8.5;
            for (int i = 9; i <= 16; i++) {
                worksheet.Columns[i].ColumnWidth = 5;
            }

            for (int i = 6; i <= 23; i++) {
                worksheet.Columns[i].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }

        }

        /// <summary>
        ///     Create the summary page to show pretty charts and figures in a concise format
        ///     this one is extermely long, due to having a hard time being able to dynamically
        ///     create the seperate charts, as needed to show a difference between dispatch/site
        ///     and kiosk types of queues.
        ///     a majority of this method is formatting, that i wish could be in another area, but
        ///     it wasn't easy/possible
        /// </summary>
        /// <param name="sites"></param>
        /// <param name="workbook"></param>
        /// <param name="remote"></param>
        private void AddSummaryPage(ref List<Site> sites, ref Excel.Workbook workbook, bool local = true) {
            List<List<ResolutionGroup>> Queues;

            if (!local) {
                Queues = fixRemoteSiteList(sites);
            } else {
                Queues = SeperateSiteList(sites);
            }

            Excel.Worksheet currentSheet = workbook.Worksheets.Add();
            currentSheet.Name = "CSAT Summary";
            excelApp.Windows.Application.ActiveWindow.DisplayGridlines = false;
            currentSheet.Cells.WrapText = true;

            //formatting
            currentSheet.Columns[1].ColumnWidth = 13;
            currentSheet.Columns[2].ColumnWidth = 9.5;
            currentSheet.Columns[3].ColumnWidth = 9.5;
            currentSheet.Columns[4].ColumnWidth = 9.5;
            currentSheet.Columns[5].ColumnWidth = 9.5;
            currentSheet.Columns[6].ColumnWidth = 1.5;

            //grabs the date from the hidden "Date" sheet previously created
            currentSheet.Cells[2, 1] = "=Date!A1";
            currentSheet.Range[currentSheet.Cells[2, 1], currentSheet.Cells[2, 20]].MergeCells = true;
            currentSheet.Rows[2].RowHeight = 25;
            currentSheet.Cells[2, 1].NumberFormat = "[$-409]mmmm yyyy;@";
            currentSheet.Cells[2, 1].Font.Size = 24;
            currentSheet.Cells[2, 1].Font.Bold = true;
            currentSheet.Cells[2, 1].Font.Color = Excel.XlRgbColor.rgbBlue;
            currentSheet.Cells[2, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            currentSheet.Cells[2, 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //these made working with the chart sizes and placing them on the summary page a lot easier to deal with.
            int lastRow;
            int chartT = 60;
            int chartL = 325;
            int chartW = 312;
            int chartH = 175;
            foreach (List<ResolutionGroup> rgList in Queues) {
                if (rgList.Count > 0) {
                    //capture the last row being used, to dynamically build the following page, this is set repeatedly.
                    lastRow = getMaxRow(currentSheet) + 2;
                    currentSheet.Cells[lastRow, 1] = rgList[0].Type + " Customer Satisfaction Survey Summary";
                    currentSheet.Cells[lastRow, 1].Font.Bold = true;
                    currentSheet.Cells[lastRow, 1].Font.Color = Excel.XlRgbColor.rgbBlue;
                    currentSheet.Cells[lastRow, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    currentSheet.Range[currentSheet.Cells[lastRow, 1], currentSheet.Cells[lastRow, 5]].MergeCells = true;

                    lastRow = getMaxRow(currentSheet) + 2;
                    currentSheet.Cells[lastRow, 1] = "Site";
                    currentSheet.Cells[lastRow, 2] = "Survey Responses";
                    currentSheet.Cells[lastRow, 3] = "Raw Rating";
                    currentSheet.Cells[lastRow, 4] = "Filtered Responses";
                    currentSheet.Cells[lastRow, 5] = "Filtered Rating";
                    currentSheet.Range[currentSheet.Cells[lastRow, 1], currentSheet.Cells[lastRow, 5]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    currentSheet.Range[currentSheet.Cells[lastRow, 1], currentSheet.Cells[lastRow, 5]].Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = Excel.XlRgbColor.rgbBlue;
                    currentSheet.Rows[lastRow].RowHeight = 39;
                    currentSheet.Rows[lastRow].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    currentSheet.Rows[lastRow].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                    lastRow = getMaxRow(currentSheet) + 1;
                    int tableTop = lastRow;

                    foreach (ResolutionGroup group in rgList) {
                        if (group.CSATS.Count > 0) {
                            currentSheet.Cells[lastRow, 1] = group.Name;
                            currentSheet.Cells[lastRow, 2] = "=" + group.Name + "!C4";
                            currentSheet.Cells[lastRow, 3] = "=" + group.Name + "!C2";
                            currentSheet.Cells[lastRow, 4] = "=" + group.Name + "!C5";
                            currentSheet.Cells[lastRow, 5] = "=" + group.Name + "!C3";
                            currentSheet.Rows[lastRow].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        }
                        if (rgList.IndexOf(group) == (rgList.Count - 1)) {

                        }
                        lastRow = getMaxRow(currentSheet) + 1;
                    }
                    int tableBottom = lastRow - 1;

                    currentSheet.Range[currentSheet.Cells[tableBottom, 1], currentSheet.Cells[tableBottom, 5]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    currentSheet.Range[currentSheet.Cells[tableBottom, 1], currentSheet.Cells[tableBottom, 5]].Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = Excel.XlRgbColor.rgbBlue;

                    lastRow = getMaxRow(currentSheet) + 1;
                    currentSheet.Rows[lastRow].Font.Bold = true;
                    currentSheet.Rows[lastRow].Font.Color = Excel.XlRgbColor.rgbBlue;
                    currentSheet.Rows[lastRow].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    currentSheet.Rows[lastRow].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    currentSheet.Cells[lastRow, 1] = "AVERAGE";
                    currentSheet.Cells[lastRow, 2] = "=AVERAGE(B" + tableTop.ToString() + ":B" + tableBottom.ToString() + ")";
                    currentSheet.Cells[lastRow, 3] = "=AVERAGE(C" + tableTop.ToString() + ":C" + tableBottom.ToString() + ")";
                    currentSheet.Cells[lastRow, 4] = "=AVERAGE(D" + tableTop.ToString() + ":D" + tableBottom.ToString() + ")";
                    currentSheet.Cells[lastRow, 5] = "=AVERAGEIF(E" + tableTop.ToString() + ":E" + tableBottom.ToString() + ", \">0\")";
                    currentSheet.Range[currentSheet.Cells[lastRow, 2], currentSheet.Cells[lastRow, 2]].NumberFormat = "0.00";

                    lastRow = getMaxRow(currentSheet) + 1;
                    int totalRow = lastRow;
                    currentSheet.Rows[lastRow].Font.Bold = true;
                    currentSheet.Rows[lastRow].Font.Color = Excel.XlRgbColor.rgbBlue;
                    currentSheet.Rows[lastRow].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    currentSheet.Rows[lastRow].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    currentSheet.Cells[lastRow, 1] = "TOTAL";
                    currentSheet.Cells[lastRow, 2] = "=SUM(B" + tableTop.ToString() + ":B" + tableBottom.ToString() + ")";
                    currentSheet.Cells[lastRow, 4] = "=SUM(D" + tableTop.ToString() + ":D" + tableBottom.ToString() + ")";
                    string sheetsForRawAvg = "";
                    string sheetsForFilterAvg = "";
                    for (int i = 0; i <= (tableBottom - tableTop); i++) {
                        sheetsForRawAvg += "INDIRECT(" + ColumnNumToString(1) + (tableTop + i).ToString() + " & \"!C7\")";
                        sheetsForFilterAvg += "INDIRECT(" + ColumnNumToString(1) + (tableTop + i).ToString() + " & \"!C8\")";
                        if (i != (tableBottom - tableTop)) {
                            sheetsForRawAvg += " + ";
                            sheetsForFilterAvg += " + ";
                        }
                    }
                    lastRow = getMaxRow(currentSheet) + 1;
                    currentSheet.Rows[lastRow].Font.Bold = true;
                    currentSheet.Rows[lastRow].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    currentSheet.Rows[lastRow].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    currentSheet.Cells[lastRow, 2] = "Raw Avg:";
                    currentSheet.Cells[lastRow, 3] = "=(" + sheetsForRawAvg + ") / B" + totalRow.ToString() + "";
                    currentSheet.Cells[lastRow, 3].NumberFormat = "0.00";
                    currentSheet.Cells[lastRow, 4] = "Filter Avg:";
                    currentSheet.Cells[lastRow, 5] = "=(" + sheetsForFilterAvg + ") / (B" + totalRow.ToString() + " - D" + totalRow.ToString() + ")";
                    currentSheet.Cells[lastRow, 5].NumberFormat = "0.00";

                    lastRow = getMaxRow(currentSheet) + 1;
                    currentSheet.Rows[lastRow].Font.Color = Excel.XlRgbColor.rgbRed;
                    currentSheet.Rows[lastRow].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    currentSheet.Rows[lastRow].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    currentSheet.Cells[lastRow, 2] = "Target:  4.51 or 90.2%";
                    currentSheet.Cells[lastRow, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    currentSheet.Range[currentSheet.Cells[lastRow, 2], currentSheet.Cells[lastRow, 5]].MergeCells = true;

                    //time to make the pretty bar charts
                    Excel.ChartObjects xlCharts = (Excel.ChartObjects)currentSheet.ChartObjects();
                    //get the range of names for the chart
                    Excel.Range categoryRange = currentSheet.Range[currentSheet.Cells[tableTop, 1], currentSheet.Cells[tableBottom, 1]];

                    //collect the expected data for the chart
                    Excel.Range dataRange = currentSheet.Range[currentSheet.Cells[tableTop, 3], currentSheet.Cells[tableBottom, 3]];
                    //create the first chart
                    Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(chartL, chartT, chartW, chartH);
                    Excel.Chart chart = BuildChart(myChart, dataRange, categoryRange);
                    chart.ChartTitle.Text = rgList[0].Type + " Raw Ratings";

                    //rinse and repeat the above.
                    dataRange = currentSheet.Range[currentSheet.Cells[tableTop, 5], currentSheet.Cells[tableBottom, 5]];
                    myChart = (Excel.ChartObject)xlCharts.Add(chartL + chartW, chartT, chartW, chartH);
                    chart = BuildChart(myChart, dataRange, categoryRange);
                    chart.ChartTitle.Text = rgList[0].Type + " Filtered Ratings";

                    //move the chartT[op] variable to the next position, if another set of charts is to be created.
                    chartT += 235;
                }
            }
        }

        /// <summary>
        ///     turn the list of sites into a single, sorted List to be used in the summary page
        /// </summary>
        /// <param name="sites"></param>
        /// <returns></returns>
        private List<List<ResolutionGroup>> SeperateSiteList(List<Site> sites) {
            var allQs = new List<List<ResolutionGroup>>();
            var kQ = new List<ResolutionGroup>();
            var dQ = new List<ResolutionGroup>();
            foreach (Site site in sites) {
                foreach (ResolutionGroup group in site.Groups) {
                    switch (group.Type) {
                        case "Kiosk":
                            kQ.Add(group);
                            break;
                        case "Dispatch":
                        case "Site":
                            dQ.Add(group);
                            break;
                    }
                }
            }
            //reverse the List, so it's back to the original order
            kQ.Reverse();
            allQs.Add(kQ);

            dQ.Reverse();
            allQs.Add(dQ);

            return allQs;
        }

        /// <summary>
        ///     turn the list of sites into a single List to be used in the summary page
        /// </summary>
        /// <param name="sites"></param>
        /// <returns></returns>
        private List<List<ResolutionGroup>> fixRemoteSiteList(List<Site> sites) {
            var allQs = new List<List<ResolutionGroup>>();
            var Qs = new List<ResolutionGroup>();
            foreach(Site site in sites) {
                foreach (ResolutionGroup group in site.Groups) {
                    Qs.Add(group);
                }
            }
            Qs.Reverse();
            allQs.Add(Qs);

            return allQs;
        }

        /// <summary>
        ///     builds the pretty chart and returns it
        /// </summary>
        /// <param name="chartObj">  the ChartObject created by the other method </param>
        /// <param name="data"> 
        ///     this is the range of cells with data for the chart
        /// </param>
        /// <param name="categories">
        ///     this is the range of cells with the categories for the chart
        /// </param>
        /// <returns>
        ///     a lovely chart with all the fixin's
        /// </returns>
        private Excel.Chart BuildChart(Excel.ChartObject chartObj, Excel.Range data, Excel.Range categories) {
            Excel.Chart chart = chartObj.Chart;
            //set the source data for the chart
            chart.SetSourceData(data);
            chart.HasLegend = false;
            chart.HasTitle = true;
            chart.ChartTitle.Font.Bold = false;

            //set the categories for the chart
            Excel.Axis xAxis = chart.Axes(Excel.XlAxisType.xlValue);
            Excel.Axis yAxis = chart.Axes(Excel.XlAxisType.xlCategory);
            Excel.TickLabels ticks = xAxis.TickLabels;
            xAxis.MinimumScale = 3;
            xAxis.MaximumScale = 5;
            xAxis.MajorTickMark = Excel.XlTickMark.xlTickMarkOutside;
            xAxis.MinorTickMark = Excel.XlTickMark.xlTickMarkInside;
            xAxis.MajorUnit = 0.5;
            xAxis.MinorUnit = 0.1;
            xAxis.HasMinorGridlines = false;
            xAxis.HasMajorGridlines = true;
            yAxis.CategoryNames = categories;

            //set up pretty colors on the bars
            Excel.ChartGroup chartGroup = (Excel.ChartGroup)chart.ChartGroups(1);
            chartGroup.VaryByCategories = true;
            return chart;
        }

        /// <summary>
        ///     returns the month name when given a month number
        /// </summary>
        /// <param name="monthNum"> number of the month </param>
        /// <param name="ShortMonthName"> 
        ///     when this is true, the method will return a shortened version of the month name 
        /// </param>
        /// <returns></returns>
        static private string getMonthName(int monthNum, bool ShortMonthName = false) {
            string month = "";
            switch (monthNum) {
                case 1:
                    month = "January";
                    break;
                case 2:
                    month = "February";
                    break;
                case 3:
                    month = "March";
                    break;
                case 4:
                    month = "April";
                    break;
                case 5:
                    month = "May";
                    break;
                case 6:
                    month = "June";
                    break;
                case 7:
                    month = "July";
                    break;
                case 8:
                    month = "August";
                    break;
                case 9:
                    month = "September";
                    break;
                case 10:
                    month = "October";
                    break;
                case 11:
                    month = "November";
                    break;
                case 12:
                    month = "December";
                    break;
            }
            if (ShortMonthName && (month.Length > 3)) {
                return month.Remove(3);
            }
            return month;
        }

        /// <summary>
        ///     Opens the folder and selects the last workbook created by the application,
        ///     then proceeds to close the application.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOpenFolder_Click(object sender, EventArgs e) {
            if (FileFolder != null) {
                string argument = "/select, \"" + FilePath + "\"";
                System.Diagnostics.Process.Start("explorer.exe", argument);
                Close();
            }
        }

        /// <summary>
        ///     dispose of everything and close the application window
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClose_Click(object sender, EventArgs e) {
            Close();
        }

        /// <summary>
        ///     when the app is closing, make sure to close out of the opened excel application window
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CSATFixer_Closing(object sender, EventArgs e) {
            excelApp.Quit();
        }

        /// <summary>
        ///     detect if the selection has changed in the data grid view and set the checkbox accordingly
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvLocalSites_SelectionChanged(object sender, EventArgs e) {
            DataGridView sent = sender as DataGridView;
            if (!sent.AreAllCellsSelected(true)) {
                cbSelectAllLocalSites.Checked = false;
            }            
        }

        /// <summary>
        ///     detect if the selection has changed in the data grid view and set the checkbox accordingly
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvRemoteSites_SelectionChanged(object sender, EventArgs e) {
            DataGridView sent = sender as DataGridView;
            if (!sent.AreAllCellsSelected(true)) {
                cbSelectAllRemoteSites.Checked = false;
            }
        }

        /// <summary>
        ///     the SelectAllSites checkbox logic to make sure it either selects or deselects all objects,
        /// depending on the state of the checkmark
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbSelectAllLocalSites_CheckedChanged(object sender, EventArgs e) {
            CheckBox sent = sender as CheckBox;
            if (sent.Checked) {
                dgvLocalSites.SelectAll();
            }
        }

        /// <summary>
        ///     the SelectAllRemoteSites checkbox logic to make sure it either selects or deselects all objects,
        /// depending on the state of the checkmark
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbSelectAllRemoteSites_CheckedChanged(object sender, EventArgs e) {
            CheckBox sent = sender as CheckBox;
            if (sent.Checked) {
                dgvRemoteSites.SelectAll();
            }
        }
    }

    /// <summary>
    ///     The individual CSAT object that will hold all of the required properties for building the worksheets
    ///     I have added the : IEquatable interface, so that a comparison can be made to prevent duplicates from being stored.
    /// </summary>
    public class CSAT : IEquatable<CSAT>{
        public string TicketID { get; set; }
        public string ResolutionSpecialist { get; set; }
        public string Originator { get; set; }
        public int Q1Score { get; set; }
        public int Q2Score { get; set; }
        public int Q3Score { get; set; }
        public int Q4Score { get; set; }
        public int Q5Score { get; set; }
        public int Q6Score { get; set; }
        public int Q7Score { get; set; }
        public int Q8Score { get; set; }
        public string AssignmentGroup { get; set; }
        public string CustomerComment { get; set; }
        private const int NumScores = 8;

        public CSAT() {

        }
        
        private double averageScore(bool includeZero = false) {
            if (!includeZero) {
                return Math.Round(((Q1Score + Q2Score + Q3Score + Q4Score + Q5Score + Q6Score + Q7Score + Q8Score) / NumNonZeroScores()), 2);
            }
            return Math.Round(((Q1Score + Q2Score + Q3Score + Q4Score + Q5Score + Q6Score + Q7Score + Q8Score) / (double)NumScores), 2);
        }

        /// <summary>
        ///     finds how many are not 0, due to how the average is calculated in the worksheet
        /// </summary>
        /// <returns>
        ///     number of Q scores that are non-zero
        /// </returns>
        private double NumNonZeroScores() {
            double count = 0.0;
            if (Q1Score > 0)
                count++;
            if (Q2Score > 0)
                count++;
            if (Q3Score > 0)
                count++;
            if (Q4Score > 0)
                count++;
            if (Q5Score > 0)
                count++;
            if (Q6Score > 0)
                count++;
            if (Q7Score > 0)
                count++;
            if (Q8Score > 0)
                count++;
            return count;
        }
        
        public bool AboveTarget(double target) {
            return (averageScore() > target);
        }

        // the logic required to be able to compare CSATs to each other
        public override bool Equals(Object obj) {
            if (obj == null) {
                return false;
            }
            CSAT objAsCSAT = obj as CSAT;
            if (objAsCSAT == null) {
                return false;
            } else {
                return Equals(objAsCSAT);
            }
        }

        public override int GetHashCode() {
            return (Q1Score + Q2Score + Q3Score + Q4Score + Q5Score + Q6Score + Q7Score + Q8Score).GetHashCode();
        }

        public bool Equals(CSAT other) {
            if (other == null) {
                return false;
            }
            return (this.TicketID.Equals(other.TicketID));
        }

        public static bool operator ==(CSAT lhs, CSAT rhs) {
            if (object.ReferenceEquals(lhs, null)) {
                return object.ReferenceEquals(rhs, null);
            }
            return lhs.Equals(rhs);
        }

        public static bool operator !=(CSAT lhs, CSAT rhs) {
            if (object.ReferenceEquals(lhs, null)) {
                return object.ReferenceEquals(rhs, null);
            }
            return !(lhs.Equals(rhs));
        }
    }

    /// <summary>
    ///     ResolutionGroup Object to hold the top-most properties of each queue of each site.
    ///     also contains the list of CSATs associated with the queue
    /// </summary>
    public class ResolutionGroup {

        public string Name { get; set; }
        public string Type { get; set; }
        public List<CSAT> CSATS = new List<CSAT>();

        public ResolutionGroup(string name, string type) {
            Name = name;
            Type = type;
        }

        public bool isEmpty() {
            return CSATS.Count == 0;
        }
    }
    
    /// <summary>
    ///     Site Object to hold the top-level properties of the individual sites
    /// </summary>
    public class Site {

        public string Name { get; set; }
        public List<ResolutionGroup> Groups = new List<ResolutionGroup>();

        public int Queues {
            get {
                return Groups.Count;
            }
        }
        
        public int CSATS {
            get {
                int count = 0;
                foreach (ResolutionGroup rg in Groups)
                    count += rg.CSATS.Count;
                return count;
            }
        }

        public Site(string siteName) {
            Name = siteName;
        }

        public void addCSAT(string groupName, CSAT csat) {
            foreach (ResolutionGroup rg in Groups) {
                if (rg.Name == groupName && !rg.CSATS.Contains(csat)) {
                    rg.CSATS.Add(csat);
                }
            }
        }

        public bool hasGroup (string groupName) {
            return Groups.Contains(findGroup(groupName));
        }

        public ResolutionGroup findGroup(string groupName) {
            foreach (ResolutionGroup rg in Groups) {
                if (groupName == rg.Name) {
                    return rg;
                }
            }
            return null;
        }

        public int getGroupIndexByName(string groupName) {
            int index = 0;
            foreach (ResolutionGroup rg in Groups) {
                if (rg.Name == groupName) {
                    break;
                }
                index++;
            }
            return index;
        }

        public int getGroupIndex(ResolutionGroup group) {
            int index = 0;
            foreach (ResolutionGroup rg in Groups) {
                if (rg.Name == Groups[index].Name) {
                    break;
                }
                index++;
            }
            return index;
        }
    }
}