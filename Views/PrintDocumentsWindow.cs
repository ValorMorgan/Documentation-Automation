using FB.AmericaMe.UI.Helpers.Documents;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FB.AmericaMe.UI.Views
{
    public partial class PrintDocumentsWindow : Form
    {
        #region PROPERTIES
        private string template;
        private BackgroundWorker worker = new BackgroundWorker() { WorkerSupportsCancellation = true };
        /// <summary>
        /// Gets the background worker thread for this window.  Cannot be set.
        /// </summary>
        public BackgroundWorker Worker
        {
            get
            {
                return worker;
            }
            set { }
        }
        #endregion

        #region CONSTRUCTOR
        public PrintDocumentsWindow()
        {
            InitializeComponent();
            worker.WorkerReportsProgress = true;
            worker.DoWork += new DoWorkEventHandler(worker_DoWork);
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            radioButtonWord.Checked = true;
            radioButtonExcel.Checked = false;
        }
        #endregion

        #region METHODS
        private void ProcessTemplate(object sender, EventArgs e)
        {
            if (sender.ToString().Contains("Only 1st"))
                template = template + "1";
            else if (sender.ToString().Contains("1st and 2nd"))
                template = template + "2";
            else
                template = template + "3";

            runCreateDocument(template);
        }

        private void EditTemplate(object sender, EventArgs e)
        {
            if (sender.ToString().Contains("Only 1st"))
                template = template + "1";
            else if (sender.ToString().Contains("1st and 2nd"))
                template = template + "2";
            else
                template = template + "3";

            MenuEdit_Click(null, null);
        }

        /// <summary>
        /// Creates a document based off the template alone.
        /// </summary>
        /// <param name="template">The template of the document to create.</param>
        private void runCreateDocument(string template)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                UpdateControls(false);
                object[] parameters = new object[] { template };
                worker.RunWorkerAsync(parameters);
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("No Data"))
                    MessageBox.Show("No data was loaded.");
                else
                    MessageBox.Show(string.Format("Exception Caught: {0}{1}{1}{2}", ex.Message, Environment.NewLine, ex.StackTrace), "Error Dialog", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Creates a LABELS document based off the template and lable type.
        /// </summary>
        /// <param name="template">The template of the document to create.</param>
        /// <param name="lbltype">The label type of the document.</param>
        private void runCreateDocument(string template, DocumentCreationHelper.enumLabeltype lbltype)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                UpdateControls(false);
                object[] parameters = new object[] { template, lbltype };
                worker.RunWorkerAsync(parameters);
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("No Data"))
                    MessageBox.Show("No data was loaded.");
                else
                    MessageBox.Show(string.Format("Exception Caught: {0}{1}{1}{2}", ex.Message, Environment.NewLine, ex.StackTrace), "Error Dialog", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Creates a JUDGE LABEL document based off the template, label type, provided judges, and starting label.
        /// </summary>
        /// <param name="template">The template of the document to create.</param>
        /// <param name="lbltype">The label type of the document.</param>
        /// <param name="judges">The list of judges to create labels for.</param>
        /// <param name="startLabel">The label to start at.</param>
        public void runCreateDocument(string template, DocumentCreationHelper.enumLabeltype lbltype, IList<string> judges = null, int startLabel = 1)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                UpdateControls(false);
                object[] parameters = new object[] { template, lbltype, judges, startLabel };
                worker.RunWorkerAsync(parameters);
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("No Data"))
                    MessageBox.Show("No data was loaded.");
                else
                    MessageBox.Show(string.Format("Exception Caught: {0}{1}{1}{2}", ex.Message, Environment.NewLine, ex.StackTrace), "Error Dialog", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Creates a CERTIFICATE document based off the template, certificate type, and implacement.
        /// </summary>
        /// <param name="template">The template of the document to create.</param>
        /// <param name="certType">The certificate type of the document.</param>
        /// <param name="inplacement">The implacement of the certificate.</param>
        private void runCreateDocument(string template, DocumentCreationHelper.CertificateType certType, string inplacement)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                UpdateControls(false);
                object[] parameters = new object[] { template, certType, inplacement };
                worker.RunWorkerAsync(parameters);
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("No Data"))
                    MessageBox.Show("No data was loaded.");
                else
                    MessageBox.Show(string.Format("Exception Caught: {0}{1}{1}{2}", ex.Message, Environment.NewLine, ex.StackTrace), "Error Dialog", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Gets the proper drawing coordinates for the clicked link
        /// based on the location it is drawn in.
        /// </summary>
        /// <param name="selectedTab">The name of the current tab that is selected.</param>
        /// <param name="s">The LinkLabel that was clicked.</param>
        /// <returns>The coordinates as a Point for drawing the element.</returns>
        private System.Drawing.Point GetMenuCoordinates(string selectedTab, LinkLabel s)
        {
            System.Drawing.Point rtnPoint = new System.Drawing.Point();

            switch (selectedTab)
            {
                case "Mailing Labels":
                    rtnPoint.X = roundedGroupBox1.Location.X + roundedGroupBox6.Location.X + ((s.Parent.Name.Contains("panel")) ? panel1.Location.X : 0) + s.Location.X + s.Margin.Left;
                    rtnPoint.Y = tabControl1.ItemSize.Height + roundedGroupBox1.Location.Y + roundedGroupBox6.Location.Y + ((s.Parent.Name.Contains("panel")) ? panel1.Location.Y : 0) + s.Location.Y + s.Height + s.Height / 2;
                    break;
                case "Memo's":
                    rtnPoint.X = roundedGroupBox3.Location.X + roundedGroupBox4.Location.X + s.Location.X + s.Margin.Left;
                    rtnPoint.Y = tabControl1.ItemSize.Height + roundedGroupBox3.Location.Y + (s.Name.Contains("Agent") ? roundedGroupBox4.Location.Y : roundedGroupBox5.Location.Y) + s.Location.Y + s.Height + s.Height / 2;
                    break;
                case "Certificates":
                    rtnPoint.X = roundedGroupBox2.Location.X + s.Location.X + s.Margin.Left;
                    rtnPoint.Y = tabControl1.ItemSize.Height + roundedGroupBox2.Location.Y + s.Location.Y + s.Height + s.Height / 2;
                    break;
                case "State Winners":
                    rtnPoint.X = roundedGroupBox7.Location.X + roundedGroupBox8.Location.X + s.Location.X + s.Margin.Left;
                    rtnPoint.Y = tabControl1.ItemSize.Height + roundedGroupBox7.Location.Y + roundedGroupBox8.Location.Y + s.Location.Y + s.Height + s.Height / 2;
                    break;
                default:
                    rtnPoint.X = 20;
                    rtnPoint.Y = 20;
                    break;
            }

            return rtnPoint;
        }

        /// <summary>
        /// Updates the controls in the window to either disable them during a process
        /// or enable them after the process is complete.
        /// </summary>
        /// <param name="enable">Whether to enable the controls or disable them.</param>
        private void UpdateControls(bool enable)
        {
            roundedGroupBox4.Enabled = enable;
            roundedGroupBox5.Enabled = enable;
            roundedGroupBox6.Enabled = enable;
            roundedGroupBox8.Enabled = enable;
            Close1Link.Enabled = enable;
            Close2Link.Enabled = enable;
            Close3Link.Enabled = enable;
            Close4Link.Enabled = enable;
            NonSponsoredCertificate123.Enabled = enable;
            SponsoredCertificate123.Enabled = enable;
            cancelButton1.Enabled = (enable) ? false : true;
            cancelButton2.Enabled = (enable) ? false : true;
            cancelButton3.Enabled = (enable) ? false : true;
            cancelButton4.Enabled = (enable) ? false : true;

            // If controls are being enabled, no process is running.
            if (enable)
                UpdateProgressBars(0);
        }
        #endregion

        #region EVENTS

        #region Label Link
        private void Label_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var s = sender as LinkLabel;
            string senderName = s.Name;
            if (senderName.Contains("Judges"))
                template = "JudgeLabels";
            else
                template = "Labels";

            /////////////////////////////////
            //         Right Click         //
            /////////////////////////////////
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                MenuItem[] menuItems2 = new MenuItem[] { new MenuItem("Edit Label Template", new System.EventHandler(this.MenuEdit_Click)) };
                ContextMenu buttonMenu2 = new ContextMenu(menuItems2);
                buttonMenu2.Show(this, GetMenuCoordinates(tabControl1.SelectedTab.Text, s));
                return;
            }

            /////////////////////////////////
            //         Left Click          //
            /////////////////////////////////
            if (senderName.Contains("School_Winners")) //School Winners
            {
                MenuItem[] schoolmenuItems = new MenuItem[]
                { 
                    new MenuItem("Print School Winners - Only 1st Place Winner", new System.EventHandler(ProcessLabel)),
                    new MenuItem("Print School Winners - 1st and 2nd Place Winners", new System.EventHandler(ProcessLabel)),
                    new MenuItem("Print School Winners - 1st, 2nd and 3rd Place Winners", new System.EventHandler(ProcessLabel)),
                };

                ContextMenu schoolbuttonMenu = new ContextMenu(schoolmenuItems);
                schoolbuttonMenu.Show(this, GetMenuCoordinates(tabControl1.SelectedTab.Text, s));
            }
            else if (senderName.Contains("Agent_Winners")) //Agent Winners
            {
                MenuItem[] agentmenuItems = new MenuItem[]
                { 
                    new MenuItem("Print Agent Winners - Only 1st Place Winner", new System.EventHandler(ProcessLabel)),
                    new MenuItem("Print Agent Winners - 1st and 2nd Place Winners", new System.EventHandler(ProcessLabel)),
                    new MenuItem("Print Agent Winners - 1st, 2nd and 3rd Place Winners", new System.EventHandler(ProcessLabel)),
                };

                ContextMenu agentbuttonMenu = new ContextMenu(agentmenuItems);
                agentbuttonMenu.Show(this, GetMenuCoordinates(tabControl1.SelectedTab.Text, s));
            }
            else if (senderName.Contains("AllSchools"))
            {
                this.runCreateDocument(template, DocumentCreationHelper.enumLabeltype.allschool);
            }
            else if (senderName.Contains("AllParticipating"))
            {
                this.runCreateDocument(template, DocumentCreationHelper.enumLabeltype.allparticipants);
            }
            else if (senderName.Contains("Sponsors"))
            {
                this.runCreateDocument(template, DocumentCreationHelper.enumLabeltype.sponsors);
            }
            else if (senderName.Contains("Judges"))
            {
                Views.JudgesLabelPrintWindow frm = new Views.JudgesLabelPrintWindow();
                frm.saveFormat = (radioButtonWord.Checked) ? "Word (.doc)" : "Excel (.xlsx)";
                frm.PopulateJudgesListBox();
                if (DialogResult.Cancel != frm.ShowDialog())
                {
                    this.runCreateDocument(template, DocumentCreationHelper.enumLabeltype.judges, frm.judges, Convert.ToInt32(frm.StartAtLabelNumericUpDown.Value));
                }
                frm.Dispose();
            }
        }

        /// <summary>
        /// Click-Event Handler to help process links in one location.
        /// </summary>
        /// <param name="sender">The link clicked.</param>
        /// <param name="e"></param>
        private void ProcessLabel(object sender, EventArgs e)
        {
            MenuItem item = (MenuItem)sender;

            // Filters parts of the sender.
            // ie Sender: "Print Agent Winners - 1st and 2nd Place Winners"
            string senderHead = (item.Text.Contains('-')) ?
                item.Text.Split('-')[0].Trim().Replace("Print ", "") : item.Text.Trim().Replace("Print ", "");
            string senderTail = (item.Text.Contains('-')) ?
                item.Text.Split('-')[1].Trim() : item.Text.Trim();

            try
            {
                switch (senderHead)
                {
                    case "School Winners":
                        if (senderTail.Contains("3"))
                            this.runCreateDocument(template, DocumentCreationHelper.enumLabeltype.thirdplaceSchoolLabel);
                        else if (senderTail.Contains("2"))
                            this.runCreateDocument(template, DocumentCreationHelper.enumLabeltype.secondplaceSchoolLabel);
                        else if (senderTail.Contains("1"))
                            this.runCreateDocument(template, DocumentCreationHelper.enumLabeltype.firstplaceSchoolLabel);
                        else
                            throw new Exception("ERROR: Failed to parse SenderTail in ProcessLabel properly!");
                        break;

                    case "Agent Winners":
                        if (senderTail.Contains("3"))
                            this.runCreateDocument(template, DocumentCreationHelper.enumLabeltype.thirdplaceAgentLabel);
                        else if (senderTail.Contains("2"))
                            this.runCreateDocument(template, DocumentCreationHelper.enumLabeltype.secondplaceAgentLabel);
                        else if (senderTail.Contains("1"))
                            this.runCreateDocument(template, DocumentCreationHelper.enumLabeltype.firstplaceAgentLabel);
                        else
                            throw new Exception("ERROR: Failed to parse SenderTail in ProcessLabel properly!");
                        break;

                    default:
                        throw new NotImplementedException("ProcessLabel needs more functionality.  SenderHead: " + senderHead);
                }
            }
            catch (NotImplementedException ex)
            {
                MessageBox.Show(ex.Message + "\n\nPlease contact a developer for assistance.", "ProcessLabel needs more functionality", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n\nPlease contact a developer for assistance.", "ProcessLabel Failur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void radioButtonWord_Clicked(object sender, EventArgs e)
        {
            radioButtonWord.Checked = true;
            radioButtonExcel.Checked = false;
        }

        private void radioButtonExcel_Clicked(object sender, EventArgs e)
        {
            radioButtonWord.Checked = false;
            radioButtonExcel.Checked = true;
        }
        #endregion

        #region Document Link
        private void Document_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var s = sender as LinkLabel;
            template = s.Name;
            try
            {
                /////////////////////////////////
                //         Right Click         //
                /////////////////////////////////
                if (e.Button == System.Windows.Forms.MouseButtons.Right)
                {
                    MenuItem[] menuItems;
                    if (template.Contains("123"))//@=Template has three files attached to it
                    {
                        template = string.Format(template.Replace("123", ""), 1);

                        menuItems = new MenuItem[] {
                            new MenuItem(string.Format("Edit {0} - Only 1st Place Winner Template", template), new System.EventHandler(this.EditTemplate)),
                            new MenuItem(string.Format("Edit {0} - 1st and 2nd Place Winner Template", template), new System.EventHandler(this.EditTemplate)),
                            new MenuItem(string.Format("Edit {0} - 1st, 2nd, and 3rd Place Winner Template", template), new System.EventHandler(this.EditTemplate)) 
                        };
                    }
                    else
                    {
                        menuItems = new MenuItem[] { 
                            new MenuItem("Edit Template", new System.EventHandler(this.MenuEdit_Click)) 
                        };
                    }
                    ContextMenu buttonMenu = new ContextMenu(menuItems);
                    buttonMenu.Show(this, GetMenuCoordinates(tabControl1.SelectedTab.Text, s));
                    return;
                }


                /////////////////////////////////
                //         Left Click         //
                /////////////////////////////////
                if (template.Contains("123"))//123=Template has three files attached to it
                {
                    //Documents
                    template = template.Replace("123", "");
                    MenuItem[] menuItems = new MenuItem[]
                    { 
                        new MenuItem(string.Format("Print {0} - Only 1st Place Winner", template), new System.EventHandler(this.ProcessTemplate)),
                        new MenuItem(string.Format("Print {0} - 1st and 2nd Place Winners", template), new System.EventHandler(this.ProcessTemplate)),
                        new MenuItem(string.Format("Print {0} - 1st, 2nd and 3rd Place Winners", template), new System.EventHandler(this.ProcessTemplate)),
                    };

                    ContextMenu buttonMenu = new ContextMenu(menuItems);
                    buttonMenu.Show(this, GetMenuCoordinates(tabControl1.SelectedTab.Text, s));
                    return;

                }

                this.runCreateDocument(template);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Exception Caught: {0}{1}{1}{2}", ex.Message, Environment.NewLine, ex.StackTrace), "Error Dialog", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }
        #endregion

        #region Certificate Link
        private void NonSponsoredCertificate123_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            template = "Certificate";
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                MenuItem[] menuItems2 = new MenuItem[] { new MenuItem("Edit Label Template", new System.EventHandler(this.MenuEdit_Click)) };
                ContextMenu buttonMenu2 = new ContextMenu(menuItems2);
                buttonMenu2.Show(this, GetMenuCoordinates(tabControl1.SelectedTab.Text, (LinkLabel)sender));
                return;
            }
            MenuItem[] menuItems = new MenuItem[]
            { 
                new MenuItem(string.Format("Print Non-Sponsored Winner {0} - Only 1st Place Winner", template), new System.EventHandler(this.ProcessWinnerCertificate)),
                new MenuItem(string.Format("Print Non-Sponsored Winner {0} - 1st and 2nd Place Winners", template), new System.EventHandler(this.ProcessWinnerCertificate)),
                new MenuItem(string.Format("Print Non-Sponsored Winner {0} - 1st, 2nd and 3rd Place Winners", template), new System.EventHandler(this.ProcessWinnerCertificate)),
            };

            ContextMenu buttonMenu = new ContextMenu(menuItems);
            buttonMenu.Show(this, GetMenuCoordinates(tabControl1.SelectedTab.Text, (LinkLabel)sender));
        }

        private void SponsoredCertificate123_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            template = "Certificate";
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                MenuItem[] menuItems2 = new MenuItem[] { new MenuItem("Edit Label Template", new System.EventHandler(this.MenuEdit_Click)) };
                ContextMenu buttonMenu2 = new ContextMenu(menuItems2);
                buttonMenu2.Show(this, GetMenuCoordinates(tabControl1.SelectedTab.Text, (LinkLabel)sender));
                return;
            }
            MenuItem[] menuItems = new MenuItem[]
            { 
                new MenuItem(string.Format("Print Sponsored Winner {0} - Only 1st Place Winner", template), new System.EventHandler(this.ProcessWinnerCertificate)),
                new MenuItem(string.Format("Print Sponsored Winner {0} - 1st and 2nd Place Winners", template), new System.EventHandler(this.ProcessWinnerCertificate)),
                new MenuItem(string.Format("Print Sponsored Winner {0} - 1st, 2nd and 3rd Place Winners", template), new System.EventHandler(this.ProcessWinnerCertificate)),
            };

            ContextMenu buttonMenu = new ContextMenu(menuItems);
            buttonMenu.Show(this, GetMenuCoordinates(tabControl1.SelectedTab.Text, (LinkLabel)sender));
        }

        private void ProcessWinnerCertificate(object sender, EventArgs e)
        {
            DocumentCreationHelper.CertificateType certType;
            if (sender.ToString().Contains("Non-"))
                certType = DocumentCreationHelper.CertificateType.NonSponsored;
            else
                certType = DocumentCreationHelper.CertificateType.Sponsored;

            template = string.Format("{0}_{1}", template, certType.ToString());

            if (sender.ToString().Contains("Only 1st"))
                runCreateDocument(template + "1", certType, DocumentCreationHelper.placement.first);
            else if (sender.ToString().Contains("1st and 2nd"))
                runCreateDocument(template + "2", certType, DocumentCreationHelper.placement.second);
            else
                runCreateDocument(template + "3", certType, DocumentCreationHelper.placement.third);
        }
        #endregion

        void CloseWindow_LinkClicked(object o, LinkLabelLinkClickedEventArgs e)
        {
            this.Close();
        }

        private void MenuEdit_Click(object sender, EventArgs e)
        {
            //  LinkLabel s = sender as LinkLabel;
            Helpers.Documents.DocumentCreationHelper.OpenAsWordDocument(Helpers.Documents.DocumentRepositoryHelper.GetDocumentTemplate(template));
        }
        #endregion

        #region ProgressBar
        /// <summary>
        /// The starting time for the progress bar.
        /// </summary>
        private DateTime startTime;
        /// <summary>
        /// How much time is remaining that is used for the estimated time.
        /// </summary>
        private TimeSpan timeRemaining = TimeSpan.Zero;
        /// <summary>
        /// If the first estimate needs to be made.
        /// </summary>
        private bool initialEstimate = true;

        /// <summary>
        /// Updates all the progress bars to the current process progress.
        /// </summary>
        /// <param name="progress">The current amount of progress.</param>
        private void UpdateProgressBars(int progress)
        {
            progressBar_Labels.Value = progress;
            progressBar_Memos.Value = progress;
            progressBar_Certificates.Value = progress;
            progressBar_Winners.Value = progress;
        }

        /// <summary>
        /// Updates the progress bar with the proper amount.  A check to invoke
        /// exists to determine if the main thread needs to process anything.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void worker_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            // Prevent no progress or reverse progress (only need to check one progress bar)
            if (e.ProgressPercentage <= progressBar_Labels.Value)
                return;

            // Progress has started, start the estimateTimer
            if (e.ProgressPercentage == 1)
            {
                estimateTimer.Start();
                startTime = DateTime.Now;
            }
            else if (e.ProgressPercentage > 1)
            {
                // Get the first estimate of time
                if (initialEstimate)
                {
                    // timeNow - startingTime * Total Percent
                    timeRemaining = TimeSpan.FromTicks(DateTime.Now.Subtract(startTime).Ticks * 100);
                    initialEstimate = false;
                }
                // Update the estimate every 25% and at 90% at the very end.
                else if (e.ProgressPercentage % 25 == 0 || e.ProgressPercentage == 90 || e.ProgressPercentage >= 98)
                {
                    // timeNow - startingTime * ((Total Percent - Current Percent) / Current Percent) * 2
                    // Converts exist to allow proper math
                    // Multiply by 2 for more accuracy
                    timeRemaining = TimeSpan.FromTicks((int)((double)(DateTime.Now.Subtract(startTime).Ticks)
                        * ((double)(100 - e.ProgressPercentage) / (double)e.ProgressPercentage))
                        * 2);
                }
            }
            
            // Update the progress bar.  Checks exist to keep a range of 0-100
            if (this.InvokeRequired)
            {
                this.Invoke(new MethodInvoker(delegate
                {
                    UpdateProgressBars((e.ProgressPercentage > 100) ? 100 : (e.ProgressPercentage < 0) ? 0 : e.ProgressPercentage);
                }));
            }
            else
            {
                UpdateProgressBars((e.ProgressPercentage > 100) ? 100 : (e.ProgressPercentage < 0) ? 0 : e.ProgressPercentage);
            }
        }

        /// <summary>
        /// The work of the worker.  A check on the parameters is used to determine the
        /// document to create.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            estimateTimer.Interval = 1000;
            timeRemaining = TimeSpan.Zero;
            initialEstimate = true;
            string saveFormat = "";
            this.Invoke(new MethodInvoker(delegate
            {
                saveFormat = (radioButtonWord.Checked) ? "Word (.doc)" : "Excel (.xlsx)";
            }));

            try
            {
                object[] parameters = e.Argument as object[];
                if (parameters.Count() <= 1)
                    DocumentCreationHelper.CreateDocument(parameters[0].ToString(), worker);
                else if (parameters.Count() == 2)
                    DocumentCreationHelper.ProcessMailLabels(parameters[0].ToString(), (DocumentCreationHelper.enumLabeltype)parameters[1], saveFormat, worker);
                else if (parameters.Count() == 3)
                {
                    Type type = parameters[2].GetType();
                    if (type == typeof(string))
                        Helpers.Documents.DocumentCreationHelper.CreateDocument(parameters[0].ToString(), worker, (int)parameters[1], parameters[2].ToString());
                    else
                        DocumentCreationHelper.ProcessMailLabels(parameters[0].ToString(), (DocumentCreationHelper.enumLabeltype)parameters[1], saveFormat, worker, (List<string>)parameters[2]);
                }
                else if (parameters.Count() == 4)
                {
                    Type type = parameters[2].GetType();
                    if (type == typeof(string))
                        Helpers.Documents.DocumentCreationHelper.CreateDocument(parameters[0].ToString(), worker, (int)parameters[1], parameters[2].ToString(), (int)parameters[3]);
                    else
                        DocumentCreationHelper.ProcessMailLabels(parameters[0].ToString(), (DocumentCreationHelper.enumLabeltype)parameters[1], saveFormat, worker, (List<string>)parameters[2], (int)parameters[3]);
                }
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("No Data"))
                    MessageBox.Show("No data was loaded.");
                else
                    MessageBox.Show(string.Format("Exception Caught: {0}{1}{1}{2}", ex.Message, Environment.NewLine, ex.StackTrace), "Error Dialog", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Operations that are run after worker_DoWork completes.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            UpdateControls(true);
            Cursor.Current = Cursors.Default;
            estimateTimer.Stop();
            updateTimerLabels();
            worker.Dispose();
        }

        /// <summary>
        /// Handles the cancel request from the user.
        /// </summary>
        private void worker_CancelRequest()
        {
            estimateTimer.Stop();
            updateTimerLabels();
            worker.CancelAsync();
            worker.Dispose();
        }

        /// <summary>
        /// When the cancel button is pressed, the worker is told to stop.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cancelButton_Click(object sender, EventArgs e)
        {
            worker_CancelRequest();
        }
        private void cancelButton2_Click(object sender, EventArgs e)
        {
            worker_CancelRequest();
        }
        private void cancelButton3_Click(object sender, EventArgs e)
        {
            worker_CancelRequest();
        }
        private void cancelButton4_Click(object sender, EventArgs e)
        {
            worker_CancelRequest();
        }

        /// <summary>
        /// Timer event to update the timer labels with proper estimate.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void estimateTimer_Tick(object sender, EventArgs e)
        {
            // Get the time as a string
            int minutes = timeRemaining.Minutes;
            int seconds = timeRemaining.Seconds;
            string time = "";
            if (minutes > 0 || seconds > 0)
            {
                time = string.Format("{0}:{1}", ((minutes.ToString().Length <= 1) ? string.Format("0{0}", minutes) : minutes.ToString())
                                              , ((seconds.ToString().Length <= 1) ? string.Format("0{0}", seconds) : seconds.ToString()));
            }

            // Update the estime label with the time
            updateTimerLabels(time);

            // Subtract a second and make sure to not go negative
            if (timeRemaining.Seconds + (timeRemaining.Minutes * 60) != 0)
                timeRemaining = timeRemaining.Subtract(TimeSpan.FromSeconds(1));
        }

        /// <summary>
        /// Updates the timerLabels with the proper time.
        /// If called with no parameters, the label reads "Estimated Time: ".
        /// </summary>
        /// <param name="time">The amount of time to put in the label.</param>
        private void updateTimerLabels(string time = "")
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new MethodInvoker(delegate
                {
                    timerLabel1.Text = string.Format("Estimated Time: {0}", time);
                    timerLabel2.Text = string.Format("Estimated Time: {0}", time);
                    timerLabel3.Text = string.Format("Estimated Time: {0}", time);
                    timerLabel4.Text = string.Format("Estimated Time: {0}", time);
                }));
            }
            else
            {
                timerLabel1.Text = string.Format("Estimated Time: {0}", time);
                timerLabel2.Text = string.Format("Estimated Time: {0}", time);
                timerLabel3.Text = string.Format("Estimated Time: {0}", time);
                timerLabel4.Text = string.Format("Estimated Time: {0}", time);
            }
        }
        #endregion
    }
}