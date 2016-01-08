using FB.AmericaMe.Model;
using FB.Common.Extensions;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
namespace FB.AmericaMe.UI.Helpers.Documents
{
    public class DocumentCreationHelper
    {
        #region DOCUMENTS
        internal static void CreateDocument(string template, System.ComponentModel.BackgroundWorker worker, int certType = -1, string inplacement = "", int startLabel = 1)
        {
            try
            {
                DocumentRepositoryHelper.DeleteOldFiles();
                worker.ReportProgress(0); // Work started
                ProcessDocument(template, worker, certType, inplacement, startLabel);
                worker.ReportProgress(100); // Work finished
            }
            catch (FileLoadException)
            {
                /* Cancel request from user */
                worker.ReportProgress(0);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void ProcessDocument(string template, System.ComponentModel.BackgroundWorker worker, int certType = -1, string inplacement = "", int startLabel = 1)
        {
            //Create the file first            
            string argTempFileName = Helpers.Documents.DocumentRepositoryHelper.TempDocumentFileName(template);
            RichTextBox rtf = new RTFExtended();
            bool createFile = true;

            try
            {
                // If the same file already exists, ask to use that one
                string latestFile = DocumentRepositoryHelper.GetLatestFile(template);
                if (latestFile != null)
                {
                    DialogResult choice = MessageBox.Show(string.Format("This document was created on {0}.  Open that one?"
                        , System.IO.File.GetCreationTime(latestFile))
                        , "Recent Document Found"
                        , MessageBoxButtons.YesNoCancel);
                    if (choice == DialogResult.Yes)
                    {
                        createFile = false;
                        argTempFileName = latestFile;
                    }
                    else if (choice == DialogResult.Cancel)
                        throw new FileLoadException();
                }

                if (createFile)
                {
                    switch (template.Split('_')[0])
                    {
                        case DocumentTemplates.Certificate:
                            rtf.Rtf = BuildCertificateRTF(template, getCertificateData(template, certType, inplacement), worker);
                            rtf.SaveFile(argTempFileName, RichTextBoxStreamType.RichNoOleObjs);
                            break;
                        case DocumentTemplates.TopTenTeacherLetter:
                        case DocumentTemplates.TopTenStudentLetter:
                        case DocumentTemplates.TopTenNewsRelease:
                            BuildWordDocument<TopTenDocumentModel>(template, null, null, worker);
                            break;
                        default:
                            BuildWordDocument<DocumentModel>(template, null, null, worker);
                            break;
                    };
                }

                if (System.IO.File.Exists(argTempFileName))
                    OpenAsWordDocument(argTempFileName);
                else
                    throw new Exception("No Data");
            }
            catch (FileLoadException)
            {
                /* Cancel request from user */
                throw;
            }
            catch (Exception)
            {
                throw;
            }
        }

        #region RTF
        public static string BuildRTF(string template, IList<DocumentModel> looplist, System.ComponentModel.BackgroundWorker worker = null)
        {
            string templatelocation = Helpers.Documents.DocumentRepositoryHelper.GetDocumentTemplate(template.Split('_')[0]);
            StringBuilder sbRtf = new StringBuilder();
            StringBuilder sbRtfTemp = new StringBuilder();
            RichTextBox TemplateHolder = new RTFExtended();

            TemplateHolder.LoadFile(templatelocation);
            if (looplist.Count <= 0)
                throw new Exception("No Data Found.");

            foreach (var item in looplist)
            {
                sbRtfTemp.Append(TemplateHolder.Rtf.Substring(1, TemplateHolder.Rtf.LastIndexOf('}') - 1));

                sbRtfTemp = sbRtfTemp.Replace(DocumentPlaceHolders.agentname, item.AgentName);
                sbRtfTemp = sbRtfTemp.Replace(DocumentPlaceHolders.agentfname, item.AgentFirstName);
                sbRtfTemp = sbRtfTemp.Replace(DocumentPlaceHolders.agentcity, item.AgentCity);
                sbRtfTemp = sbRtfTemp.Replace(DocumentPlaceHolders.date, DateTime.Now.ToString("MMM, yyyy"));
                sbRtfTemp = sbRtfTemp.Replace(DocumentPlaceHolders.schoolname, item.SchoolName);
                sbRtfTemp = sbRtfTemp.Replace(DocumentPlaceHolders.schoolcity, item.SchoolCity);
                sbRtfTemp = sbRtfTemp.Replace(DocumentPlaceHolders.teachername, item.TeacherName);
                sbRtfTemp = sbRtfTemp.Replace(DocumentPlaceHolders.contesttopic, item.ContestTopic);
                sbRtfTemp = sbRtfTemp.Replace(DocumentPlaceHolders.contestyearspan, item.ContestYearSpan);
                sbRtfTemp = sbRtfTemp.Replace(DocumentPlaceHolders.firstplacestudent, item.FirstPlaceStudent);
                sbRtfTemp = sbRtfTemp.Replace(DocumentPlaceHolders.firstplacestudentfname, item.FirstPlaceStudentFName);
                sbRtfTemp = sbRtfTemp.Replace(DocumentPlaceHolders.his_her, item.HisHer);
                sbRtfTemp = sbRtfTemp.Replace(DocumentPlaceHolders.secondplacestudent, item.SecondPlaceStudent);
                sbRtfTemp = sbRtfTemp.Replace(DocumentPlaceHolders.thirdplacestudent, item.ThirdPlaceStudent);

                sbRtf.Append(sbRtfTemp);
                if (looplist.Last() == item)
                    sbRtf.Append(@"\rtf1");
                else
                    sbRtf.Append(@"\rtf1 \page");
                sbRtfTemp.Clear();

                if (worker != null)
                {
                    if (worker.CancellationPending)
                    {
                        throw new ApplicationException();
                    }

                    double progress = ((double)(looplist.IndexOf(item) + 1) / (double)looplist.Count) * 100.00;
                    worker.ReportProgress((int)progress);
                }
            }

            return string.Format("{0}{1}{2}", "{", sbRtf.ToString(), "}");
        }
        public static string BuildCertificateRTF(string template, IList<CertificateModel> looplist, System.ComponentModel.BackgroundWorker worker = null)
        {
            string templatelocation = Helpers.Documents.DocumentRepositoryHelper.GetDocumentTemplate(template.Split('_')[0]);
            StringBuilder sbRtf = new StringBuilder();
            StringBuilder sbRtfTemp = new StringBuilder();
            RichTextBox TemplateHolder = new RTFExtended();

                    TemplateHolder.LoadFile(templatelocation);

            if (looplist.Count <= 0)
                MessageBox.Show("No data was collected or found.");

            foreach (var item in looplist)
            {
                        sbRtfTemp.Append(TemplateHolder.Rtf.Substring(1, TemplateHolder.Rtf.LastIndexOf('}') - 1));

                sbRtfTemp = sbRtfTemp.Replace(CertificatePlaceHolder.StudentName, item.StudentName);
                sbRtfTemp = sbRtfTemp.Replace(CertificatePlaceHolder.Placement, item.Placement.ToString());
                sbRtfTemp = sbRtfTemp.Replace(CertificatePlaceHolder.SchoolName, item.SchoolName);

                sbRtf.Append(sbRtfTemp);
                sbRtf.Append(@"\rtf1 \par \page");
                sbRtfTemp.Clear();

                if (worker != null)
                {
                    if (worker.CancellationPending)
                    {
                        throw new ApplicationException();
                    }

                    double progress = ((double)(looplist.IndexOf(item) + 1) / (double)looplist.Count) * 100.00;
                    worker.ReportProgress((int)progress);
                }
            }
            return string.Format("{0}{1}{2}", "{", sbRtf.ToString(), "}");
        }
        #endregion
        #region WORD
        public static void BuildWordDocument<T>(string template, List<T> records = null, List<String> arguments = null, System.ComponentModel.BackgroundWorker worker = null)
        {
            // Get the template location and filename
            string templatelocation = Helpers.Documents.DocumentRepositoryHelper.GetDocumentTemplate(template.Split('_')[0]);
            object filename = Helpers.Documents.DocumentRepositoryHelper.TempDocumentFileName(template);

            IList<DocumentModel> looplistDocumentModel = new List<DocumentModel>();
            IList<TopTenDocumentModel> looplistTopTenDocumentModel = new List<TopTenDocumentModel>();
            if (records == null)
            {
                switch(typeof(T).Name)
                {
                    case "DocumentModel":
                        looplistDocumentModel = getDocumentData(template);
                        if (looplistDocumentModel.Count <= 0)
                            throw new Exception("No Data Found.");
                        break;

                    case "TopTenDocumentModel":
                        looplistTopTenDocumentModel = getTopTenData(template);
                        if (looplistTopTenDocumentModel.Count <= 0)
                            throw new Exception("No Data Found.");
                        break;

                    default:
                        throw new NotImplementedException("BuildWordDocument needs more functionality.\nUnable to get data for type " + typeof(T).Name + ".");
                }
            }
            
            bool workerPresent = (worker == null) ? false : true;

            Word.Application winword = null;
            Word.Document document = null;

            try
            {
                // Create an instance of the Word application
                winword = new Word.Application();

                // Set status for word application to be visible or not.
                winword.Visible = false;

                // Create a missing variable for missing value
                object missing = System.Reflection.Missing.Value;

                // Create a template variable for the document template
                object docTemplate = templatelocation;

                // Create a blank new document and select it
                document = winword.Documents.Add(ref docTemplate, ref missing, ref missing, false);
                document.Select();

                // Create selection parameter objects
                object gotoPage = Word.WdGoToItem.wdGoToPage;
                object gobyUnits = Word.WdUnits.wdCharacter;
                object breakPage = Word.WdBreakType.wdPageBreak;
                int charactersPerPage = document.Characters.Count;//document.ComputeStatistics(Word.WdStatistic.wdStatisticCharactersWithSpaces, false);

                // Arguments Variables
                int recordStart = 1;
                string type = null;

                // Parsing arguments
                if (arguments != null)
                {
                    for (int i = 0; i < arguments.Count; i++)
                    {
                        if (arguments[i].Contains("startLabel"))
                        {
                            recordStart = Convert.ToInt32(arguments[i].Substring(11));
                            if (recordStart < 1)
                                recordStart = 1;
                        }
                        else
                        {
                            switch (arguments[i].ToLower())
                            {
                                case "labels":
                                    type = "labels";
                                    break;

                                default:
                                    throw new NotImplementedException("BuildWordDocument needs more functionality.\nArgument '" + arguments[i] + "' not known.");
                            }
                        }
                    }
                }

                // Start the estimate timer
                if (workerPresent)
                    worker.ReportProgress(1);

                Word.Range range;

                if (records == null)
                {
                    #region looplist
                    if (looplistDocumentModel.Count > 0)
                    {
                        #region looplistDocumentModel
                        foreach (var item in looplistDocumentModel)
                        {
                            if (workerPresent && worker.CancellationPending)
                            {
                                ((Word._Document)document).Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                                ((Word._Application)winword).Quit(Word.WdSaveOptions.wdDoNotSaveChanges);

                                throw new FileLoadException();
                            }

                            //Move selection to whole page
                            winword.Selection.MoveEnd(ref gobyUnits, charactersPerPage);

                            range = winword.Selection.Range;

                            //Move insertion point to end of page
                            winword.Selection.MoveStart(ref gobyUnits, charactersPerPage);

                            if (looplistDocumentModel.IndexOf(item) > 0)
                            {
                                //Break to the next page
                                winword.Selection.InsertBreak(ref breakPage);

                                //Insert next template page
                                winword.Selection.InsertFile(templatelocation);
                            }

                            #region Replace Holders
                            //Multi-threaded procedure using 4 tasks and having the system wait for
                            //all tasks to complete.
                            var tasks = new List<System.Threading.Tasks.Task>() { };
                            tasks.Add(System.Threading.Tasks.Task.Factory.StartNew(delegate
                            {
                                range.Find.Execute(DocumentPlaceHolders.agentname, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.AgentName, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(DocumentPlaceHolders.agentfname, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.AgentFirstName, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(DocumentPlaceHolders.agentcity, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.AgentCity, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(DocumentPlaceHolders.date, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   DateTime.Now.ToString("MMM, yyyy"), Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                            }));
                            tasks.Add(System.Threading.Tasks.Task.Factory.StartNew(delegate
                            {
                                range.Find.Execute(DocumentPlaceHolders.schoolname, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.SchoolName, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(DocumentPlaceHolders.schoolcity, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.SchoolCity, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(DocumentPlaceHolders.teachername, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.TeacherName, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(DocumentPlaceHolders.contesttopic, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.ContestTopic, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                            }));
                            tasks.Add(System.Threading.Tasks.Task.Factory.StartNew(delegate
                            {
                                range.Find.Execute(DocumentPlaceHolders.contestyearspan, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.ContestYearSpan, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(DocumentPlaceHolders.firstplacestudent, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.FirstPlaceStudent, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(DocumentPlaceHolders.firstplacestudentfname, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.FirstPlaceStudentFName, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(DocumentPlaceHolders.his_her, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.HisHer, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                            }));
                            tasks.Add(System.Threading.Tasks.Task.Factory.StartNew(delegate
                            {
                                range.Find.Execute(DocumentPlaceHolders.secondplacestudent, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.SecondPlaceStudent, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(DocumentPlaceHolders.thirdplacestudent, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.ThirdPlaceStudent, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                            }));

                            System.Threading.Tasks.Task.WaitAll(tasks.ToArray());
                            #endregion

                            if (workerPresent)
                            {
                                int progress = (int)(((double)(looplistDocumentModel.IndexOf(item) + 1) / (double)looplistDocumentModel.Count) * 100.00);
                                if (progress > 1)
                                    worker.ReportProgress(progress);
                            }
                        }
                        #endregion
                    }
                    else if (looplistTopTenDocumentModel.Count > 0)
                    {
                        #region looplistTopTenDocumentModel
                        foreach (var item in looplistTopTenDocumentModel)
                        {
                            if (workerPresent && worker.CancellationPending)
                            {
                                ((Word._Document)document).Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                                document = null;
                                ((Word._Application)winword).Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
                                winword = null;

                                // Release the objects
                                releaseObject(document);
                                releaseObject(winword);

                                throw new FileLoadException();
                            }

                            //Move selection to whole page
                            winword.Selection.MoveEnd(ref gobyUnits, charactersPerPage);

                            range = winword.Selection.Range;

                            //Move insertion point to end of page
                            winword.Selection.MoveStart(ref gobyUnits, charactersPerPage);

                            if (looplistTopTenDocumentModel.IndexOf(item) > 0)
                            {
                                //Break to the next page
                                winword.Selection.InsertBreak(ref breakPage);

                                //Insert next template page
                                winword.Selection.InsertFile(templatelocation);
                            }

                            #region Replace Holders
                            //Multi-threaded procedure using 4 tasks and having the system wait for
                            //all tasks to complete.
                            var tasks = new List<System.Threading.Tasks.Task>() { };
                            tasks.Add(System.Threading.Tasks.Task.Factory.StartNew(delegate
                            {
                                range.Find.Execute(TopTenDocumentPlaceHolders.date, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   DateTime.Now.Date.ToString("MM/dd/yyyy"), Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(TopTenDocumentPlaceHolders.toptenstudentfirstname, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.toptenstudentfirstname, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(TopTenDocumentPlaceHolders.toptenstudentlastname, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.toptenstudentlastname, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(TopTenDocumentPlaceHolders.studentaddress, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.studentaddress, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                            }));
                            tasks.Add(System.Threading.Tasks.Task.Factory.StartNew(delegate
                            {
                                range.Find.Execute(TopTenDocumentPlaceHolders.studentcity, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.studentcity, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(TopTenDocumentPlaceHolders.studentzip, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.studentzip, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(TopTenDocumentPlaceHolders.him_her, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.him_her, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(TopTenDocumentPlaceHolders.parentnames, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.parentnames, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                            }));
                            tasks.Add(System.Threading.Tasks.Task.Factory.StartNew(delegate
                            {
                                range.Find.Execute(TopTenDocumentPlaceHolders.parentcity, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.parentcity, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(TopTenDocumentPlaceHolders.son_daughter, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.son_daughter, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(TopTenDocumentPlaceHolders.teacher, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.teachername, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(TopTenDocumentPlaceHolders.schoolname, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.schoolname, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                            }));
                            tasks.Add(System.Threading.Tasks.Task.Factory.StartNew(delegate
                            {
                                range.Find.Execute(TopTenDocumentPlaceHolders.schooladdress, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.schooladdress, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(DocumentPlaceHolders.schoolcity, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.schoolcity, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(TopTenDocumentPlaceHolders.schoolzip, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.schoolzip, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(DocumentPlaceHolders.contesttopic, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.contesttopic, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                            }));
                            tasks.Add(System.Threading.Tasks.Task.Factory.StartNew(delegate
                            {
                                range.Find.Execute(DocumentPlaceHolders.contestyearspan, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.contestyearspan, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(TopTenDocumentPlaceHolders.placement, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                   item.placement, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                            }));

                            System.Threading.Tasks.Task.WaitAll(tasks.ToArray());

                            if (template == DocumentTemplates.TopTenNewsRelease)
                            {
                                var toptenwinners = getTopTenWinners();
                                for (int i = 0; i <= 10; i++)
                                {
                                    var winner = toptenwinners.Where(w => w.Participant.StateWinnerPlacement == i);
                                    if (winner.Count() > 0)
                                    {
                                        range.Find.Execute(TopTenDocumentPlaceHolders.toptenstudent.Replace("#", i.ToString()), ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                           winner.First().Student.Person.FullName, Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                        range.Find.Execute(TopTenDocumentPlaceHolders.toptenschool.Replace("#", i.ToString()), ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                           string.Format("{0}, {1}", winner.First().Participant.School.Name, winner.First().Participant.School.Address.City), Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                    }
                                }
                            }
                            #endregion

                            if (workerPresent)
                            {
                                int progress = (int)(((double)(looplistTopTenDocumentModel.IndexOf(item) + 1) / (double)looplistTopTenDocumentModel.Count) * 100.00);
                                if (progress > 1)
                                    worker.ReportProgress(progress);
                            }
                        }

                        #endregion
                    }
                    #endregion
                }
                else
                {
                    #region records
                    //Move selection to whole page
                    winword.Selection.MoveEnd(ref gobyUnits, charactersPerPage);

                    range = winword.Selection.Range;

                    //Move insertion point to end of page
                    winword.Selection.MoveStart(ref gobyUnits, charactersPerPage);

                    // Determines the number of elements on a page.
                    // Generally used with argument "labels"
                    // WARNING: Regex will grab ALL numbers regardless of being
                    //          inside brackets or not.
                    int elementsPerPage = 1;
                    if (type.Equals("labels"))
                    {
                        System.Text.RegularExpressions.MatchCollection numbers = System.Text.RegularExpressions.Regex.Matches(range.Text, @"\d+");
                        elementsPerPage = Convert.ToInt32(numbers[numbers.Count - 1].Value);
                    }

                    foreach (T record in records)
                    {
                        int recordNumber = (records.IndexOf(record) % (elementsPerPage + 1)) + 1;

                        // Move to a designated record start
                        if ((records.IndexOf(record) + 1) < recordStart)
                            continue;

                        if (workerPresent && worker.CancellationPending)
                        {
                            ((Word._Document)document).Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                            ((Word._Application)winword).Quit(Word.WdSaveOptions.wdDoNotSaveChanges);

                            throw new FileLoadException();
                        }

                        if ((records.IndexOf(record) + 1) % (elementsPerPage + 1) == 0)
                        {
                            //Break to the next page
                            winword.Selection.InsertBreak(ref breakPage);

                            //Insert next template page
                            winword.Selection.InsertFile(templatelocation);

                            //Move selection to whole page
                            winword.Selection.MoveEnd(ref gobyUnits, charactersPerPage);

                            //Labels is unique to how the range setup works
                            //Generally because the template is a table setup in the Word document
                            if (!type.Equals("labels"))
                                range = winword.Selection.Range;

                            //Move insertion point to end of page
                            winword.Selection.MoveStart(ref gobyUnits, charactersPerPage);
                        }

                        // Check the type provided in the function call to use here
                        switch (typeof(T).Name)
                        {
                            #region JUDGE
                            case "JudgeLabelModel":
                                range.Find.Execute(JudgeLabelHelper.JudgeLabelPlaceHolder.Judge.Replace("#", Convert.ToInt32(recordNumber).ToString()), ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                       record.GetType().GetProperty("JudgeName").GetValue(record, null), Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(JudgeLabelHelper.JudgeLabelPlaceHolder.SchoolName.Replace("#", Convert.ToInt32(recordNumber).ToString()), ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                       record.GetType().GetProperty("SchoolName").GetValue(record, null), Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(JudgeLabelHelper.JudgeLabelPlaceHolder.SchoolCity.Replace("#", Convert.ToInt32(recordNumber).ToString()), ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                       record.GetType().GetProperty("SchoolCity").GetValue(record, null), Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(JudgeLabelHelper.JudgeLabelPlaceHolder.NumberOfEssays.Replace("#", Convert.ToInt32(recordNumber).ToString()), ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                       record.GetType().GetProperty("NumOfEssays").GetValue(record, null), Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                range.Find.Execute(JudgeLabelHelper.JudgeLabelPlaceHolder.Date.Replace("#", Convert.ToInt32(recordNumber).ToString()), ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                       record.GetType().GetProperty("Date").GetValue(record, null), Word.WdReplace.wdReplaceAll, ref missing, ref missing, ref missing, ref missing);
                                break;
                            #endregion

                            default:
                                throw new NotImplementedException("BuildWordDocument needs more functionalities.\nData type '" + typeof(T).Name + "' not known.");
                        }

                        // Report progress to user
                        if (workerPresent)
                        {
                            int progress = (int)(((double)(records.IndexOf(record) + 1) / (double)records.Count) * 100.00);
                            if (progress > 1)
                                worker.ReportProgress(progress);
                        }
                    }
                    #endregion
                }

                // Perform any cleanup
                removeRemainingPlaceholders(winword, document, charactersPerPage);
                clearBlankPages(document);
                
                //Save the document
                document.SaveAs(ref filename, Word.WdSaveFormat.wdFormatDocument);
                ((Word._Document)document).Close(Word.WdSaveOptions.wdSaveChanges, Word.WdOriginalFormat.wdWordDocument);
                ((Word._Application)winword).Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
            }
            catch (FileLoadException)
            {
                /* Cancel request from user */
                throw;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                // Release the objects
                releaseObject(document);
                releaseObject(winword);
            }

        }
        #endregion
        #region EXCEL
        public static void BuildExcelDocument<T>(string template, List<T> records, List<string> arguments = null, System.ComponentModel.BackgroundWorker worker = null)
        {
            if (records.Count <= 0)
                throw new Exception("No data was passed in for creating an Excel document");

            // Get the template filename
            string filename = Helpers.Documents.DocumentRepositoryHelper.TempDocumentFileName(template + "_SpreadSheet", "xlsx").Split('.')[0];

            bool workerPresent = (worker == null) ? false : true;

            // Setup these variables for final release in finally
            Excel.Application winexcel = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                // Create an instance for excel app
                winexcel = new Excel.Application();

                // Set the number of sheets to be made in a new workbook
                winexcel.SheetsInNewWorkbook = 1;

                // Create a missing variable for missing value
                object missing = Type.Missing;

                // Create a blank new workbook
                workbook = winexcel.Workbooks.Add(missing);

                // Create a blank new worksheet
                worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);

                // Set status for excel application to be visible or not.
                winexcel.Visible = false;

                // Prevents any pop-ups due to conflicts
                winexcel.DisplayAlerts = false;

                // Start the estimate timer
                if (workerPresent)
                    worker.ReportProgress(1);

                // Determined mode to create workbook
                string selectedMode = null;

                // Keeps track of the current row of the spreadsheet to isnert at
                int currentRow = 0;

                // Which record to start at (1 being the first)
                int recordStart = 1;

                // Parsing arguments
                if (arguments != null)
                {
                    for (int i = 0; i < arguments.Count; i++)
                    {
                        if (arguments[i].Contains("startLabel"))
                        {
                            recordStart = Convert.ToInt32(arguments[i].Substring(11));
                            if (recordStart < 1)
                                recordStart = 1;
                        }
                        else
                        {
                            switch (arguments[i].ToLower())
                            {
                                case "company mail":
                                    selectedMode = arguments[i];
                                    currentRow++;

                                    // Setup headers
                                    worksheet.Cells[1, 1] = "Agent";
                                    worksheet.Cells[1, 2] = "Company";
                                    worksheet.Cells[1, 3] = "Address";
                                    worksheet.Cells[1, 4] = "City";
                                    worksheet.Cells[1, 5] = "State";
                                    worksheet.Cells[1, 6] = "Zip";
                                    break;

                                case "judge mail":
                                    selectedMode = arguments[i];
                                    currentRow++;

                                    // Setup headers
                                    worksheet.Cells[1, 1] = "Judge";
                                    worksheet.Cells[1, 2] = "School";
                                    worksheet.Cells[1, 3] = "Address";
                                    worksheet.Cells[1, 4] = "City";
                                    worksheet.Cells[1, 5] = "State";
                                    worksheet.Cells[1, 6] = "Zip";
                                    break;

                                case "mail":
                                    selectedMode = arguments[i];
                                    currentRow++;

                                    // Setup headers
                                    worksheet.Cells[1, 1] = "School";
                                    worksheet.Cells[1, 2] = "Address";
                                    worksheet.Cells[1, 3] = "City";
                                    worksheet.Cells[1, 4] = "State";
                                    worksheet.Cells[1, 5] = "Zip";
                                    break;

                                default:
                                    throw new NotImplementedException("BuildExcelDocument needs more functionality.\nArgument '" + arguments[i] + "' not known.");
                            }
                        }
                    }
                }

                foreach (T record in records)
                {
                    if (workerPresent && worker.CancellationPending)
                    {
                        workbook.Saved = true; // To prevent popup asking to save
                        workbook.Close(Excel.XlSaveAction.xlDoNotSaveChanges);
                        winexcel.Quit();

                        throw new FileLoadException();
                    }

                    // Skip any records that were requested to be skipped
                    if (recordStart > records.IndexOf(record) + 1)
                        continue;

                    // Updates the row insertion area
                    currentRow++;

                    if (string.IsNullOrEmpty(selectedMode))
                    {
                        switch (record.ToString().Split('.')[record.ToString().Split('.').ToList().Count - 1].ToLower())
                        {
                            default:
                                throw new NotImplementedException("BuildExcelDocument needs more functionality.\nData type '" + typeof(T).Name + "' not known.");
                        }
                    }
                    else
                    {
                        Address address;
                        Person person;
                        switch (selectedMode)
                        {
                            case "company mail":
                                address = (Address)record.GetType().GetProperty("Address").GetValue(record, null);
                                person = (Person)record.GetType().GetProperty("Person").GetValue(record, null);
                                worksheet.Cells[currentRow, 1] = person.FullName;
                                worksheet.Cells[currentRow, 2] = record.GetType().GetProperty("Agency").GetValue(record, null);
                                worksheet.Cells[currentRow, 3] = (address.Address1 + address.Address2).Trim();
                                worksheet.Cells[currentRow, 4] = address.City;
                                worksheet.Cells[currentRow, 5] = "MI";
                                worksheet.Cells[currentRow, 6] = address.Zip;
                                break;

                            case "judge mail":
                                worksheet.Cells[currentRow, 1] = record.GetType().GetProperty("JudgeName").GetValue(record, null);
                                worksheet.Cells[currentRow, 2] = record.GetType().GetProperty("SchoolName").GetValue(record, null);
                                worksheet.Cells[currentRow, 3] = record.GetType().GetProperty("SchoolAddress").GetValue(record, null);
                                worksheet.Cells[currentRow, 4] = record.GetType().GetProperty("SchoolCity").GetValue(record, null);
                                worksheet.Cells[currentRow, 5] = "MI";
                                worksheet.Cells[currentRow, 6] = record.GetType().GetProperty("SchoolZip").GetValue(record, null);
                                break;

                            case "mail":
                                address = (Address)record.GetType().GetProperty("Address").GetValue(record, null);
                                worksheet.Cells[currentRow, 1] = record.GetType().GetProperty("Name").GetValue(record, null);
                                worksheet.Cells[currentRow, 2] = (address.Address1 + address.Address2).Trim();
                                worksheet.Cells[currentRow, 3] = address.City;
                                worksheet.Cells[currentRow, 4] = "MI";
                                worksheet.Cells[currentRow, 5] = address.Zip;
                                break;
                            
                            default:
                                throw new NotImplementedException("BuildExcelDocument needs more functionality.\nMode '" + selectedMode + "' not known.");
                        }
                    }

                    if (workerPresent)
                    {
                        int progress = (int)(((double)(records.IndexOf(record) + 1) / (double)records.Count) * 100.00);
                        if (progress > 1)
                            worker.ReportProgress(progress);
                    }
                }

                // Auto-fit the column width to the contents
                worksheet.Columns.AutoFit();

                // Save the workbook
                workbook.SaveAs(filename, missing, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlNoChange,
                    Excel.XlSaveConflictResolution.xlLocalSessionChanges, missing, missing, missing, missing);
                workbook.Close(true, filename, missing);
                winexcel.Quit();
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                // Release the objects
                // *** NOTE *** Garbage Collection catches process.  Will be running while program
                // is running.  Functionalities in class do not allow a cleaner way to close the process.
                releaseObject(worksheet);
                releaseObject(workbook);
                releaseObject(winexcel);
            }
        }
        #endregion

        #region GET DATA
        private static IList<DocumentModel> getDocumentData(string template)
        {
            IList<DocumentModel> rtnList = new List<DocumentModel>();
            var contest = Helpers.ContestEntryHelper.Contest;
            var contestyearspan = string.Format("{0} - {1}", contest.StartDate.Year, contest.EndDate.Year);
            var contesttopic = contest.Topic;

            switch (template)
            {
                case DocumentTemplates.SchoolTeacherLetter1:
                case DocumentTemplates.SchoolNewsRelease1:
                    return getNewsReleaseData(enumReleaseDocType.School, 1);

                case DocumentTemplates.SchoolTeacherLetter2:
                case DocumentTemplates.SchoolNewsRelease2:
                    return getNewsReleaseData(enumReleaseDocType.School, 2);

                case DocumentTemplates.SchoolTeacherLetter3:
                case DocumentTemplates.SchoolNewsRelease3:
                    return getNewsReleaseData(enumReleaseDocType.School, 3);

                case DocumentTemplates.SchoolTeacherLetter123:
                case DocumentTemplates.SchoolNewsRelease123:
                    break;

                case DocumentTemplates.AgentMemo1:
                case DocumentTemplates.AgentTeacherLetter1:
                case DocumentTemplates.AgentNewsRelease1:
                    return getNewsReleaseData(enumReleaseDocType.Agent, 1);

                case DocumentTemplates.AgentMemo2:
                case DocumentTemplates.AgentTeacherLetter2:
                case DocumentTemplates.AgentNewsRelease2:
                    return getNewsReleaseData(enumReleaseDocType.Agent, 2);

                case DocumentTemplates.AgentMemo3:
                case DocumentTemplates.AgentTeacherLetter3:
                case DocumentTemplates.AgentNewsRelease3:
                    return getNewsReleaseData(enumReleaseDocType.Agent, 3);

                case DocumentTemplates.AgentNewsRelease123:
                case DocumentTemplates.AgentTeacherLetter123:
                    break;

                case DocumentTemplates.AgentConfirmationLetter:
                    return getAgentConfirmationData();

                case DocumentTemplates.TeacherConfirmationLetter:
                    return getTeacherConfirmationData();
            }
            return rtnList;
        }
        private static IList<DocumentModel> getTeacherConfirmationData()
        {
            var rtnList = new List<DocumentModel>();
            var educators = Helpers.MainWindowHelper.MainWindow.ContestController.GetEducatorsByContest(Helpers.ContestEntryHelper.Contest);
            int i = 0;
            foreach (var item in educators)
            {
                //only add to rtnList if there is no email attached
                //this is because we only want to print the documents
                //to those without email.  We are sending this document to 
                //those with emails.
                if (item == null)
                    continue;
                if (item.Person == null)
                    continue;
                if (!string.IsNullOrEmpty(item.Person.Email))
                    rtnList.Add(new DocumentModel
                    {
                        AgentName = "",
                        TeacherName = item.Person.FullName,
                        SchoolName = item.School.Name,
                        SchoolCity = item.School.Address.City,
                        ContestTopic = Helpers.ContestEntryHelper.Contest.Topic
                    }); i++;
            }
            return rtnList;
        }
        private static IList<DocumentModel> getAgentConfirmationData()
        {
            var rtnList = new List<DocumentModel>();
            var sponsors = Helpers.MainWindowHelper.MainWindow.SponsorController
                .GetActiveSponsorsbyContest(Helpers.ContestEntryHelper.Contest);
            foreach (var item in sponsors)
            {
                var edu = MainWindowHelper.MainWindow.EducatorController.EducatorsBySchool(item.Participant.School).FirstOrDefault();
                //only add to rtnList if there is no email attached
                if (!string.IsNullOrEmpty(edu.Person.Email))
                    rtnList.Add(new DocumentModel
                    {
                        AgentName = string.Format("{0} {1}", item.Agent.Person.FirstName, item.Agent.Person.LastName),
                        TeacherName = edu.Person.FullNameWithEmail,
                        SchoolName = item.Participant.School.Name,
                        SchoolCity = item.Participant.School.Address.City,
                        ContestTopic = Helpers.ContestEntryHelper.Contest.Topic
                    });
            }
            return rtnList;
        }
        private static IList<CertificateModel> getCertificateData(string template, int certType, string inplacement)
        {
            var rtnList = new List<CertificateModel>();
            IList<Model.Participant> looplist = null;

            switch (certType)
            {
                case (int)CertificateType.Sponsored:
                    looplist = Helpers.ParticipantEntryHelper.SponsoredParticipants();
                    break;

                case (int)CertificateType.NonSponsored:
                    looplist = Helpers.ParticipantEntryHelper.NonSponsoredParticipants();
                    break;

            }
            foreach (var participant in looplist)
            {
                IList<Entry> entries = MainWindowHelper.MainWindow.ParticipantController.GetParticipantEntries(participant).OrderBy(e => e.Placement).ToList();
                var edu = MainWindowHelper.MainWindow.EducatorController.EducatorsBySchool(participant.School).FirstOrDefault();



                if (entries.Count == 0)
                    continue;

                switch (inplacement)
                {
                    case placement.first:
                        if (entries.Count != 1)
                            continue;

                        rtnList.Add(new CertificateModel()
                        {
                            StudentName = entries.FirstOrDefault().Student.Person.FullName,
                            Placement = placement.first.ToString(),
                            SchoolName = participant.School.Name

                        });
                        break;
                    case placement.second:
                        if (entries.Count != 2)
                            continue;
                        foreach (var entry in entries)
                        {
                            rtnList.Add(new CertificateModel()
                            {
                                StudentName = entry.Student.Person.FullName,
                                Placement = (entry.Placement == 1 ? placement.first : placement.second),
                                SchoolName = participant.School.Name

                            });
                        }

                        break;
                    case placement.third:
                        if (entries.Count != 3)
                            continue;
                        foreach (var entry in entries)
                        {
                            rtnList.Add(new CertificateModel()
                            {
                                StudentName = entry.Student.Person.FullName,
                                Placement = (entry.Placement == 1 ? placement.first : (entry.Placement == 2) ? placement.second : placement.third),
                                SchoolName = participant.School.Name

                            });
                        }

                        break;
                }

            }

            return rtnList;
        }
        private static IList<TopTenDocumentModel> getTopTenData(string template)
        {
            string[] placements = { "", "1st Place", "2nd Place", "3rd Place", "4th Place", "5th Place", "6th Place", "7th Place", "8th Place", "9th Place", "10th Place" };

            var rtnList = new List<TopTenDocumentModel>();

            var contest = Helpers.ContestEntryHelper.Contest;
            var contestyearspan = string.Format("{0} - {1}", contest.StartDate.Year, contest.EndDate.Year);
            var contesttopic = contest.Topic;
            var statewinners = Helpers.MainWindowHelper.MainWindow.ContestController.GetParticipants(contest).Where(p => p.StateWinner == true).OrderBy(p2 => p2.StateWinnerPlacement).ToList();

            foreach (var winner in statewinners)
            {
                var entry = winner.Entries.Where(e => e.Placement == 1).FirstOrDefault();
                var parents = entry.Student.Parents.ToList();
                string parent = "", parentaddress="", parentzip="", parentcity = "";
                
                if (parents.Count > 0)
                {
                    parentaddress = parents.FirstOrDefault().Address.Address1;
                    parentcity = parents.FirstOrDefault().Address.City;
                    parentzip = parents.FirstOrDefault().Address.Zip;

                    foreach (var p in parents)
                    {
                        if (parent != "")
                            parent = parent + " and ";

                        parent += p.Person.FullName;
                    }
                }

                rtnList.Add(new TopTenDocumentModel()
                {
                    toptenstudentfirstname = entry.Student.Person.FirstName,
                    toptenstudentlastname = entry.Student.Person.LastName,
                    placement = placements[winner.StateWinnerPlacement],
                    parentnames = parent,
                    parentcity = parentcity,
                    son_daughter = entry.Student.Person.Gender.ToLower() == "m" ? "son" : "daughter",
                    contesttopic = contesttopic,
                    contestyearspan = contestyearspan,
                    schooladdress = entry.Participant.School.Address.Address1,
                    schoolcity = entry.Participant.School.Address.City,
                    schoolzip = entry.Participant.School.Address.Zip,
                    schoolname = entry.Participant.School.Name,
                    teachername = entry.Participant.School.Educators.FirstOrDefault().Person.FullName,
                    studentaddress = parentaddress,
                    studentcity = parentcity,
                    studentzip = parentzip,
                    him_her = entry.Participant.School.Educators.FirstOrDefault().Person.Gender.ToLower() == "m" ? pronoun.him : pronoun.her
                });

            }

            return rtnList;
        }
        private static IList<Entry> getTopTenWinners()
        {
            var statewinners = Helpers.MainWindowHelper.MainWindow.ContestController.GetParticipants(Helpers.ContestEntryHelper.Contest).Where(p => p.StateWinner == true).OrderBy(p2 => p2.StateWinnerPlacement).ToList();
            var rtnList = new List<Entry>();
            var i = 1;
            while (i <= 10)
            {
                var winner = statewinners.Where(e => e.StateWinnerPlacement == i);
                if(winner.Count()>0)
                    rtnList.Add(winner.First().Entries.Where(e => e.Placement == 1).FirstOrDefault());
                
                i++;
            }

            return rtnList;
        }
        private static IList<DocumentModel> getNewsReleaseData(enumReleaseDocType docType, int placement)
        {
            IList<DocumentModel> rtnList = new List<DocumentModel>();
            var contest = Helpers.ContestEntryHelper.Contest;
            var contestyearspan = string.Format("{0} - {1}", contest.StartDate.Year, contest.EndDate.Year);
            var contesttopic = contest.Topic;
            IList<Participant> participants = null;
            Model.Sponsor sponsor = null;


            switch (docType)
            {
                case enumReleaseDocType.Agent:
                    participants = Helpers.ParticipantEntryHelper.SponsoredParticipants().ToList();
                    break;
                case enumReleaseDocType.School:
                    participants = Helpers.ParticipantEntryHelper.NonSponsoredParticipants().ToList();
                    break;
            }

            foreach (var participant in participants)
            {
                var edu = MainWindowHelper.MainWindow.EducatorController.EducatorsBySchool(participant.School).FirstOrDefault();
                IList<Entry> entries = MainWindowHelper.MainWindow.ParticipantController.GetParticipantEntries(participant);
                sponsor = Helpers.ParticipantEntryHelper.SponsorsByParticipant(participant).FirstOrDefault();

                if (entries.Count == 0)
                    continue;
                if (entries.Count != placement)
                    continue;
                var firstplaceentry = entries.Where(entry => entry.Placement == 1).First();
                if (firstplaceentry == null)
                    continue;

                DocumentModel docModel = new DocumentModel()
                {
                    ContestTopic = contesttopic,
                    ContestYearSpan = contestyearspan,
                    AgentName = (sponsor == null) ? "" : sponsor.Agent.Person.FullName.Trim(),
                    AgentFirstName = (sponsor == null) ? "" : sponsor.Agent.Person.FirstName.Trim(),
                    AgentCity = (sponsor == null) ? "" : sponsor.Agent.Address.City,
                    TeacherName = edu.Person.FullName,
                    SchoolName = participant.School.Name,
                    SchoolCity = participant.School.Address.City,
                    FirstPlaceStudent = firstplaceentry.Student.DisplayValue.Trim(),
                    FirstPlaceStudentFName = firstplaceentry.Student.Person.FirstName,
                    HisHer = firstplaceentry.Student.Person.Gender == "" ? pronoun.their : firstplaceentry.Student.Person.Gender == "M" ? pronoun.his : pronoun.her
                };

                if (placement == 1) { rtnList.Add(docModel); continue; }


                Model.Entry secondplacentry = entries.Where(entry => entry.Placement == 2).First();
                if (secondplacentry == null)
                    continue;
                docModel.SecondPlaceStudent = secondplacentry.Student.DisplayValue.Trim();
                if (placement == 2) { rtnList.Add(docModel); continue; }

                Model.Entry thirdplaceentry = entries.Where(entry => entry.Placement == 3).First();
                if (thirdplaceentry == null)
                    continue;
                docModel.ThirdPlaceStudent = thirdplaceentry.Student.DisplayValue.Trim();
                rtnList.Add(docModel);

            }
            return rtnList;
        }
        #endregion
        #endregion

        #region MAILING LABEL
        const string _ATTENTION = "Attention: 8th Grade English Teacher";

        public static void ProcessMailLabels(string template, enumLabeltype labeltype, string SaveFormat, System.ComponentModel.BackgroundWorker worker, List<string> judges = null, int startLabel = 1)
        {
            try
            {
                DocumentRepositoryHelper.DeleteOldFiles();

                // Start of work
                worker.ReportProgress(1);

                if (SaveFormat.Contains("Excel"))
                {
                    string argTempFileName = Helpers.Documents.DocumentRepositoryHelper.TempDocumentFileName(template + "_SpreadSheet", "xlsx").Split('.')[0];
                    bool createFile = true;

                    // If the same file already exists, ask to use that one
                    /*
                    string latestFile = DocumentRepositoryHelper.GetLatestFile(template, "_SpreadSheet");
                    if (latestFile != null)
                    {
                        DialogResult choice = MessageBox.Show(string.Format("This document was created on {0}.  Open that one?"
                            , System.IO.File.GetCreationTime(latestFile))
                            , "Recent Document Found"
                            , MessageBoxButtons.YesNoCancel);
                        if (choice == DialogResult.Yes)
                        {
                            createFile = false;
                            argTempFileName = latestFile;
                        }
                        else if (choice == DialogResult.Cancel)
                            throw new FileLoadException();
                    }
                    */
                    if (createFile)
                    {
                        List<string> arguments = new List<string>() { };
                        switch (labeltype)
                        {
                            case enumLabeltype.firstplaceAgentLabel:
                            case enumLabeltype.secondplaceAgentLabel:
                            case enumLabeltype.thirdplaceAgentLabel:
                            case enumLabeltype.sponsors:
                                arguments.Add("company mail");
                                BuildExcelDocument<Agent>(template, GetListBylabeltype<Agent>(labeltype), arguments, worker);
                                break;
                            case enumLabeltype.judges:
                                arguments.Add("judge mail");
                                arguments.Add("startLabel:" + startLabel.ToString());
                                BuildExcelDocument<JudgeLabelHelper.JudgeLabelModel>(template, GetListBylabeltype<JudgeLabelHelper.JudgeLabelModel>(labeltype, judges), arguments, worker);
                                break;
                            default:
                                arguments.Add("mail");
                                BuildExcelDocument<School>(template, GetListBylabeltype<School>(labeltype), arguments, worker);
                                break;
                        }
                    }

                    // End of work
                    worker.ReportProgress(100);

                    OpenAsExcelDocument(argTempFileName);
                }
                else
                {
                    string argTempFileName = Helpers.Documents.DocumentRepositoryHelper.TempDocumentFileName(template);

                    // Judges are a special case of needing to use the word document creator (margin issues)
                    if (labeltype == enumLabeltype.judges)
                    {
                        List<JudgeLabelHelper.JudgeLabelModel> documentJudges = new List<JudgeLabelHelper.JudgeLabelModel>();
                        List<ParticipantJudge> allParticipantJudges = Helpers.MainWindowHelper.MainWindow.ParticipantJudgeController.ParticipantJudges;
                        string date = DateTime.Now.Date.ToString("MM/dd/yyyy");
                        foreach(string judge in judges)
                        {
                            List<Participant> participants = allParticipantJudges.Where(p => p.Judge.Person.FullName.Equals(judge)).Select(s => s.Participant).ToList();
                            foreach(Participant participant in participants)
                            {
                                documentJudges.Add(new JudgeLabelHelper.JudgeLabelModel()
                                {
                                    JudgeName = string.Format("Materials for {0}",judge),
                                    SchoolName = participant.School.Name,
                                    SchoolCity = participant.School.Address.City,
                                    SchoolAddress = participant.School.Address.Address1 + " " + participant.School.Address.Address2,
                                    SchoolZip = participant.School.Address.Zip,
                                    NumOfEssays = string.Format("Essays Submitted: {0}", participant.EssayCount),
                                    Date = date
                                });
                            }
                        }

                        List<string> arguments = new List<string>() { "startLabel:" + startLabel, "labels" };
                        BuildWordDocument<JudgeLabelHelper.JudgeLabelModel>(template, documentJudges, arguments, worker);
                    }
                    else
                    {
                        RichTextBox rtf = new RTFExtended();
                        string finalRtf = GetLabelsRtfToPrint(template, labeltype, worker);
                        rtf.Rtf = finalRtf;

                        rtf.SaveFile(argTempFileName, RichTextBoxStreamType.RichNoOleObjs);
                    }

                    // End of work
                    worker.ReportProgress(100);

                    OpenAsWordDocument(argTempFileName);
                }
            }
            catch (FileLoadException)
            {
                /* Cancel request from user */
                worker.ReportProgress(0);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        static string GetLabelsRtfToPrint(string template, enumLabeltype lbltype, System.ComponentModel.BackgroundWorker worker = null)
        {
            string templatelocation = Helpers.Documents.DocumentRepositoryHelper.GetDocumentTemplate(template.Split('_')[0]);
            RichTextBox TemplateHolder = new RTFExtended();

            StringBuilder sbRtf = new StringBuilder();
            StringBuilder sbRtfTemp = new StringBuilder();

            TemplateHolder.LoadFile(templatelocation, RichTextBoxStreamType.RichText);

            switch (lbltype)
            {
                case enumLabeltype.allschool:
                case enumLabeltype.allparticipants:

                    //Loop through list of schools
                    var schools = GetListBylabeltype<School>(lbltype);

                    for (int counter = 1; counter <= 10; counter++)
                    {
                        foreach (School school in schools)
                        {
                            if (counter > 10)
                                counter = 1;

                            if (counter == 1)
                            {
                                sbRtf.Append(sbRtfTemp);
                                if (sbRtf.Length > 0)
                                    sbRtfTemp.Clear();

                                sbRtfTemp.Append(TemplateHolder.Rtf.Substring(1, TemplateHolder.Rtf.LastIndexOf('}') - 7));
                            }

                            sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.Header.Replace("#", counter.ToString()), "");
                            sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.Attention.Replace("#", counter.ToString()), _ATTENTION);
                            sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.CustomerName.Replace("#", counter.ToString()), school.Name);
                            sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.StreetAddress.Replace("#", counter.ToString()), string.Format("{0} {1}", school.Address.Address1, school.Address.Address2));
                            sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.StreetAddressDetail.Replace("#", counter.ToString()), string.Format("{0}, MI {1}", school.Address.City, school.Address.Zip));
                            //  sbRtfTemp.Append(@"\rtf1 \par \page");
                            counter++;

                            if (worker != null)
                            {
                                if (worker.CancellationPending)
                                {
                                    throw new FileLoadException();
                                }

                                double progress = ((double)(schools.IndexOf(school) + 1) / (double)schools.Count) * 100.00;
                                worker.ReportProgress((int)progress);
                            }
                        }

                        //Clean up the label if there are not 10 to a page
                        schools.Clear();

                        sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.Header.Replace("#", counter.ToString()), "");
                        sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.Attention.Replace("#", counter.ToString()), "");
                        sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.CustomerName.Replace("#", counter.ToString()), "");
                        sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.StreetAddress.Replace("#", counter.ToString()), "");
                        sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.StreetAddressDetail.Replace("#", counter.ToString()), "");
                    }

                    if (sbRtfTemp.Length > 0)
                    {
                        sbRtf.Append(sbRtfTemp);
                        sbRtfTemp.Clear();
                    }
                    break;

                default:

                    var objects = GetListBylabeltype<dynamic>(lbltype);
                    for (int counter = 1; counter <= 10; counter++)
                    {
                        foreach (var obj in objects)
                        {
                            if (counter > 10)
                                counter = 1;

                            if (counter == 1)
                            {
                                sbRtf.Append(sbRtfTemp);

                                if (sbRtf.Length > 0)
                                    sbRtfTemp.Clear();

                                        sbRtfTemp.Append(TemplateHolder.Rtf.Substring(1, TemplateHolder.Rtf.LastIndexOf('}') - 7));
                            }


                            string header = "", attention = "", customername = "", streetaddress = "", streetaddressdetails = "";
                            if (obj != null)
                            {
                                if (!string.IsNullOrEmpty(obj.Header))
                                    header = obj.Header;

                                if (!string.IsNullOrEmpty(obj.Attention))
                                    attention = obj.Attention;


                                if (!string.IsNullOrEmpty(obj.CompanyName))
                                    customername = obj.CompanyName;


                                if (!string.IsNullOrEmpty(obj.Address1))
                                    streetaddress = string.Format("{0} {1}", obj.Address1, obj.Address2 ?? "");

                                streetaddressdetails = string.Format("{0}, MI {1}", obj.City ?? "", obj.Zip ?? "");
                            }
                            sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.Header.Replace("#", counter.ToString()), header);
                            sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.Attention.Replace("#", counter.ToString()), attention);
                            sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.CustomerName.Replace("#", counter.ToString()), customername);
                            sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.StreetAddress.Replace("#", counter.ToString()), streetaddress);
                            sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.StreetAddressDetail.Replace("#", counter.ToString()), streetaddressdetails);
                            //  sbRtfTemp.Append(@"\rtf1 \par \page");
                            counter++;

                            if (worker != null)
                            {
                                if (worker.CancellationPending)
                                {
                                    throw new FileLoadException();
                                }

                                double progress = ((double)(objects.IndexOf(obj) + 1) / (double)objects.Count) * 100.00;
                                worker.ReportProgress((int)progress);
                            }
                        }

                        //Clean up the label if there are not 10 to a page
                        objects.Clear();

                        sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.Header.Replace("#", counter.ToString()), "");
                        sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.Attention.Replace("#", counter.ToString()), "");
                        sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.CustomerName.Replace("#", counter.ToString()), "");
                        sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.StreetAddress.Replace("#", counter.ToString()), "");
                        sbRtfTemp = sbRtfTemp.Replace(MailLabelPlaceHolder.StreetAddressDetail.Replace("#", counter.ToString()), "");

                    }


                    if (sbRtfTemp.Length > 0)
                    {
                        sbRtf.Append(sbRtfTemp);
                        sbRtfTemp.Clear();
                    }
                    break;
            }

            return string.Format("{0}{1}{2}", "{", sbRtf.ToString(), "}");
        }

        private static List<T> GetListBylabeltype<T>(enumLabeltype inlabeltype, List<string> preExistingList = null)
        {
            switch (inlabeltype)
            {
                case enumLabeltype.allschool: // Assume T = School
                    if (typeof(T) != typeof(School) && typeof(T).Namespace.Contains("FB.AmericaMe"))
                        throw new TypeLoadException("Given label type of <" + inlabeltype.ToString() + ">, function type cast needs to be of type <School>.");
                    var allSchools = MainWindowHelper.MainWindow.SchoolController.Schools.Where(s => s.IsInactive.Equals(false) || s.IsInactive.Equals(null)).ToList().OrderBy(s => s.DisplayValue).ToList();
                    return allSchools as List<T>;

                case enumLabeltype.allparticipants: // Assume T = School
                    if (typeof(T) != typeof(School) && typeof(T).Namespace.Contains("FB.AmericaMe"))
                        throw new TypeLoadException("Given label type of <" + inlabeltype.ToString() + ">, function type cast needs to be of type <School>.");
                    var participatingSchools = MainWindowHelper.MainWindow.SchoolController.GetParticipatingSchools();
                    return participatingSchools as List<T>;

                case enumLabeltype.sponsors: // Assume T = Agent
                    if (typeof(T) != typeof(Agent) && typeof(T).Namespace.Contains("FB.AmericaMe"))
                        throw new TypeLoadException("Given label type of <" + inlabeltype.ToString() + ">, function type cast needs to be of type <Agent>.");

                    var list = Helpers.MainWindowHelper.MainWindow.SponsorController.GetActiveSponsorsbyContest(Helpers.ContestEntryHelper.Contest);
                    var agents = list.Select(x => x.Agent).Distinct();

                    List<Agent> rtnListAgent = new List<Agent>();
                    List<dynamic> rtnListGeneric = new List<dynamic>();

                    foreach (var item in agents)
                    {
                        var misqlagent = Helpers.MISQLAgentHelper.SponsorController().GetAgentByAgentNumber(item.AgentNumber);
                        Agent sponsor = new Agent();
                        if (typeof(T) == typeof(Agent))
                        {
                            sponsor.Person = misqlagent.Person;
                            sponsor.Address = misqlagent.Address;
                            sponsor.Agency = UI.Helpers.Constants.COMPANYNAME;
                            rtnListAgent.Add(sponsor);
                        }
                        else
                        {
                            rtnListGeneric.Add(new
                            {
                                Header = "",
                                Attention = item.DisplayValue.ToUpper(),
                                CompanyName = UI.Helpers.Constants.COMPANYNAME,
                                Address1 = misqlagent.Address.Address1,
                                Address2 = misqlagent.Address.Address2,
                                City = misqlagent.Address.City,
                                Zip = misqlagent.Address.Zip
                            });
                        }
                    }
                    return (typeof(T) == typeof(Agent)) ? rtnListAgent.OrderBy(x => x.Address.City).ThenBy(x => x.Person.LastName).ToList() as List<T>
                                                        : rtnListGeneric as List<T>;

                case enumLabeltype.firstplaceSchoolLabel:
                case enumLabeltype.secondplaceSchoolLabel:
                case enumLabeltype.thirdplaceSchoolLabel:
                case enumLabeltype.allschoolwinners: // Assume T = School
                    if (typeof(T) != typeof(School) && typeof(T).Namespace.Contains("FB.AmericaMe"))
                        throw new TypeLoadException("Given label type of <" + inlabeltype.ToString() + ">, function type cast needs to be of type <School>.");
                    return GetSchoolWinnerLabelData<T>(inlabeltype);

                case enumLabeltype.firstplaceAgentLabel:
                case enumLabeltype.secondplaceAgentLabel:
                case enumLabeltype.thirdplaceAgentLabel:
                    if (typeof(T) != typeof(Agent) && typeof(T).Namespace.Contains("FB.AmericaMe"))
                        throw new TypeLoadException("Given label type of <" + inlabeltype.ToString() + ">, function type cast needs to be of type <Agent>.");
                    return GetAgentWinnerLabelData<T>(inlabeltype);

                case enumLabeltype.judges:
                    return JudgeLabelHelper.getJudgeLabelModels(preExistingList) as List<T>;

                default:
                    return new List<T>();
            }
        }

        private static List<T> GetSchoolWinnerLabelData<T>(enumLabeltype lbltype)
        {
            List<Entry> loopEntries = new List<Entry>();
            List<School> rtnListSchool = new List<School>();
            List<dynamic> rtnListGeneric = new List<dynamic>();

            if (lbltype == enumLabeltype.allschoolwinners)
            {
                var participants = Helpers.ContestEntryHelper.AllParticipants().Where(p => p.School.IsInactive == false && p.Entries.Count > 0).OrderBy(p => p.School.Name).ToList();
                foreach (var p in participants)
                {
                    var winner = p.Entries.Where(p1 => p1.Placement == 1);
                    if(winner.Count()>0)
                        loopEntries.Add(winner.First());                 
                }
            }
            else
            {
                var list2 = Helpers.ParticipantEntryHelper.NonSponsoredParticipants();
                foreach (var item in list2)
                {
                    var entries = MainWindowHelper.MainWindow.ParticipantController.GetParticipantEntries(item);
                    if (entries.Count == 0)
                        continue;

                    if (lbltype == enumLabeltype.firstplaceSchoolLabel)
                    {
                        if (entries.Count > 1)
                            continue;
                        loopEntries.Add(entries.First());
                    }
                    else if (lbltype == enumLabeltype.secondplaceSchoolLabel)
                    {
                        if (entries.Count != 2)
                            continue;
                        loopEntries.Add(entries.First());
                    }
                    else if (lbltype == enumLabeltype.thirdplaceSchoolLabel)
                    {
                        if (entries.Count != 3)
                            continue;

                        loopEntries.Add(entries.First());
                    }
                }
            }

            foreach (var winner in loopEntries)
            {
                Educator educator = Helpers.MainWindowHelper.MainWindow.EducatorController.EducatorsBySchool(winner.Participant.School).FirstOrDefault();
                School school = winner.Participant.School;

                if (typeof(T) == typeof(School))
                    rtnListSchool.Add(school);
                else
                {
                    rtnListGeneric.Add(new
                    {
                        Header = "",
                        Attention = winner.Participant.School.Name,
                        CompanyName = string.Format("ATT: {0}", educator.Person.FullName),
                        Address1 = school.Address.Address1,
                        Address2 = school.Address.Address2,
                        City = school.Address.City,
                        Zip = school.Address.Zip
                    });
                }
            }
            return (typeof(T) == typeof(School)) ? rtnListSchool as List<T>
                                                : rtnListGeneric as List<T>;
        }

        private static List<T> GetAgentWinnerLabelData<T>(enumLabeltype lbltype)
        {
            List<Agent> rtnListAgent = new List<Agent>();
            List<dynamic> rtnListGeneric = new List<dynamic>();
            List<Participant> list2 = Helpers.ParticipantEntryHelper.SponsoredParticipants().ToList();
            List<Entry> loopEntries = new List<Entry>();
            foreach (var item in list2)
            {
                var entries = MainWindowHelper.MainWindow.ParticipantController.GetParticipantEntries(item);
                if (entries.Count == 0)
                    continue;
                switch (lbltype)
                {
                    case enumLabeltype.firstplaceAgentLabel:
                        if (entries.Count > 1)
                            continue;
                        loopEntries.Add(entries.First());
                        break;

                    case enumLabeltype.secondplaceAgentLabel:
                        if (entries.Count != 2)
                            continue;
                        loopEntries.Add(entries.First());

                        break;

                    case enumLabeltype.thirdplaceAgentLabel:
                        if (entries.Count != 3)
                            continue;
                        loopEntries.Add(entries.First());

                        break;
                }
            }

            foreach (var winner in loopEntries)
            {
                Educator educator = Helpers.MainWindowHelper.MainWindow.EducatorController.EducatorsBySchool(winner.Participant.School).FirstOrDefault();
                School school = winner.Participant.School;
                Sponsor sponsor = Helpers.ParticipantEntryHelper.SponsorsByParticipant(winner.Participant).FirstOrDefault();

                if (typeof(T) == typeof(Agent))
                {
                    Agent agent = sponsor.Agent;
                    agent.Agency = UI.Helpers.Constants.COMPANYNAME;
                    rtnListAgent.Add(agent);
                }
                else
                {
                    rtnListGeneric.Add(new
                    {
                        Header = string.Format("Results for {0}", school.Name.Trim()),
                        Attention = UI.Helpers.Constants.COMPANYNAME,
                        CompanyName = string.Format("ATT: {0}", sponsor.DisplayValue),
                        Address1 = sponsor.Agent.Address.Address1,
                        Address2 = sponsor.Agent.Address.Address2,
                        City = sponsor.Agent.Address.City,
                        Zip = sponsor.Agent.Address.Zip
                    });
                }
            }
            return (typeof(T) == typeof(Agent)) ? rtnListAgent as List<T>
                                                : rtnListGeneric as List<T>;
        }

        #endregion

        #region METHODS
        /// <summary>
        /// Opens the document as a Word Doc
        /// </summary>
        /// <param name="argTempFileName"></param>
        public static void OpenAsWordDocument(string argTempFileName)
        {
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(argTempFileName);

            System.Diagnostics.Process newProc = new System.Diagnostics.Process();
            newProc.StartInfo.FileName = argTempFileName;
            newProc.StartInfo.Verb = "Open";
            newProc.StartInfo.ErrorDialog = true;				// Display an error dialog if the process cannot start
            try
            {
                newProc.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        /// <summary>
        /// Opens the document as a RTF Doc
        /// </summary>
        /// <param name="argTempFileName"></param>
        public static void OpenAsRTFDocument(string argTempFileName)
        {
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(argTempFileName);

            System.Diagnostics.Process newProc = new System.Diagnostics.Process();
            newProc.StartInfo.FileName = "wordpad";
            newProc.StartInfo.Arguments = argTempFileName;
            newProc.StartInfo.Verb = "Open";
            newProc.StartInfo.ErrorDialog = true;				// Display an error dialog if the process cannot start
            try
            {
                newProc.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        /// <summary>
        /// Opens the document as an Excel Doc
        /// </summary>
        /// <param name="argTempFileName"></param>
        public static void OpenAsExcelDocument(string argTempFileName)
        {
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(argTempFileName.Split('.')[0]);

            System.Diagnostics.Process newProc = new System.Diagnostics.Process();
            newProc.StartInfo.FileName = "excel";
            newProc.StartInfo.Arguments = argTempFileName.Split('.')[0];
            newProc.StartInfo.Verb = "Open";
            newProc.StartInfo.ErrorDialog = true; // Display an error dialog if the process cannot start
            try
            {
                newProc.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Handles the converting of images to a string format for RTF.
        /// </summary>
        public class DocumentImageConverter
        {
            public static string GetBase64String(string PathToFile)
            {
                try
                {
                    MemoryStream stream = new MemoryStream();
                    string newPath = Path.Combine(Environment.CurrentDirectory, PathToFile);
                    Image img = Image.FromFile(newPath);
                    img.Save(stream, System.Drawing.Imaging.ImageFormat.Bmp);

                    byte[] bytes = stream.ToArray();

                    return BitConverter.ToString(bytes, 0).Replace("-", string.Empty);
                }
                catch
                {
                    return "";
                }
            }
        }

        /// <summary>
        /// Forces the release of memory dedicated to an object and calls the garbage collector
        /// to clear any excess memory.
        /// </summary>
        /// <param name="obj">The object to deallocate.</param>
        private static void releaseObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw new Exception("Exception occured while releasing object\n\n" + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        /// <summary>
        /// [Word Only] Removes all remaining placeholders (block of text within square
        /// brackets) from the provided document.
        /// </summary>
        /// <param name="winword">The Word application that is currently open.</param>
        /// <param name="document">The Word document that is being worked on.</param>
        /// <param name="charactersPerPage">The number of characters per page.</param>
        private static void removeRemainingPlaceholders(Word.Application winword, Word.Document document, int charactersPerPage = -1)
        {
            // Sub-method for removing all placeholders within the range provided
            #region scanRange
            var scanRange = new Func<Word.Range, int>(sRange =>
            {
                // If there is no text, return
                if (sRange.Text == null)
                    return 0;

                // Find options
                object matchCase = false;
                object matchWholeWord = false;
                object matchWildCards = false;
                object matchSoundsLike = false;
                object matchAllWordForms = false;
                object forward = true;
                object format = false;
                object matchKashida = false;
                object matchDiacritics = false;
                object matchAlefHamza = false;
                object matchControl = false;
                object replace = Word.WdReplace.wdReplaceAll;
                object wrap = Word.WdFindWrap.wdFindContinue;

                // Infinite loop to find all placeholders
                int charactersRemoved = 0;
                while (true)
                {
                    // If there is no text or placeholders left, break out
                    if (sRange.Text == null || !sRange.Text.Contains("["))
                        break;

                    string placeholder;

                    // Find the placeholder
                    placeholder = string.Format("{0}{1}{2}", "[", sRange.Text.Split('[')[1].Split(']')[0].Trim(), "]");

                    // Break if a placeholder can't be found
                    if (placeholder.Equals("[]"))
                        break;

                    // Placeholders should not have any spaces or punctuation
                    // Allows document to still use square brackets if there are any used beyond placeholders
                    if (!System.Text.RegularExpressions.Regex.IsMatch(placeholder.Substring(1, placeholder.Length - 2), @"^[a-zA-Z0-9]*$"))
                        continue;

                    // Find and Replace
                    // Soft-error if the placeholder can't be found
                    string replaceText = "";
                    if (!sRange.Find.Execute(placeholder, ref matchCase, ref matchWholeWord, ref matchWildCards, ref matchSoundsLike,
                        ref matchAllWordForms, ref forward, ref wrap, ref format,
                        replaceText, ref replace, ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl))
                    {
                        System.Console.WriteLine("Failed to find placeholder: " + placeholder);
                        break;
                    }

                    charactersRemoved += placeholder.Length;
                }
                return charactersRemoved;
            });
            #endregion

            try
            {
                object missing = Type.Missing;
                int documentCharacterCount = document.Characters.Count;

                // Move selection point to the beginning
                winword.Selection.Start = documentCharacterCount;
                winword.Selection.End = winword.Selection.Start;

                // Search every searchCharacterCount amount of characters
                // Goes from end to start of document
                int searchCharacterCount = (charactersPerPage <= 0) ? 5000 : charactersPerPage;
                do
                {
                    int position = documentCharacterCount - searchCharacterCount;
                    Word.Range range = document.Range((position < 0) ? 0 : position, documentCharacterCount);

                    // Confirm we don't split on a placeholder
                    // Checks for at least 2 elements and the first not having the closing brace
                    string checkValue = range.Text.Split(']')[0];
                    if (range.Text.Split(']').Length > 1 && !checkValue.Contains('['))
                    {
                        // Shift the insertion point to just behind the opening brace
                        // If no opening brace is found, the point won't move
                        int currentEnd = winword.Selection.End;
                        winword.Selection.MoveEndUntil('[', Word.WdConstants.wdBackward);

                        // If no opening brace is found, move to the start of the document
                        if (winword.Selection.End == currentEnd)
                            winword.Selection.End = 0;

                        range = document.Range(winword.Selection.End, documentCharacterCount);
                    }

                    scanRange(range);

                    // Set the "end" position to the start
                    documentCharacterCount = position;
                } while (documentCharacterCount > 0); // Will exit once the start of the document is reached
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// [Word Only] Removes all remaining blank pages that are detected.
        /// </summary>
        /// <param name="document">The current document being worked on.</param>
        private static void clearBlankPages(Word.Document document)
        {
            foreach (Word.Section section in document.Sections)
            {
                if (section.PageSetup.SectionStart == Microsoft.Office.Interop.Word.WdSectionStart.wdSectionNewPage)
                {
                    //section.PageSetup.SectionStart = Microsoft.Office.Interop.Word.WdSectionStart.wdSectionContinuous;
                    if (section.PageSetup.SectionStart == Microsoft.Office.Interop.Word.WdSectionStart.wdSectionNewPage)
                    {
                        Word.Range rng = section.Range;
                        object oCollapseStart = Word.WdCollapseDirection.wdCollapseStart;
                        rng.Collapse(ref oCollapseStart);
                        object oChar = Word.WdUnits.wdCharacter;
                        object oNeg1 = -1;
                        rng.MoveStart(ref oChar, ref oNeg1);
                        object missing = System.Type.Missing;
                        rng.Delete(ref missing, ref missing);
                    }
                }
            }
        }
        #endregion

        #region PLACEHOLDERS & MODELS
        public class pronoun
        {
            public const string his = "his";
            public const string her = "her";
            public const string their = "their";
            public const string him = "him";
        }

        public struct placement
        {
            public const string first = "First Place";
            public const string second = "Second Place";
            public const string third = "Third Place";
        }

        #region DOCUMENTS
        public class DocumentPlaceHolders
        {
            public const string agentname = "[agentname]";
            public const string agentfname = "[agentfname]";
            public const string agentcity = "[agentcity]";

            public const string schoolname = "[schoolname]";
            public const string schoolcity = "[schoolcity]";

            public const string teachername = "[teachername]";

            public const string firstplacestudent = "[firstplacestudent]";
            public const string firstplacestudentfname = "[firstplacestudentfname]";

            public const string secondplacestudent = "[secondplacestudent]";
            public const string thirdplacestudent = "[thirdplacestudent]";

            public const string currentmonth = "[month]";
            public const string currentyear = "[year]";

            public const string startyear = "[startyear]";
            public const string endyear = "[endyear]";
            public const string his_her = "[his/her]";

            public const string contesttopic = "[contesttopic]";
            public const string contestyearspan = "[contestyearspan]";

            public const string date = "[date]";
        }

        public class DocumentModel
        {
            public virtual string AgentName { get; set; }
            public virtual string AgentFirstName { get; set; }
            public virtual string AgentCity { get; set; }
            public virtual string TeacherName { get; set; }
            public virtual string SchoolName { get; set; }
            public virtual string SchoolCity { get; set; }
            public virtual string ContestTopic { get; set; }
            public virtual string FirstPlaceStudent { get; set; }
            public virtual string FirstPlaceStudentFName { get; set; }
            public virtual string SecondPlaceStudent { get; set; }
            public virtual string SecondPlaceStudentFName { get; set; }
            public virtual string ThirdPlaceStudent { get; set; }
            public virtual string ThirdPlaceStudentFName { get; set; }
            public virtual string HisHer { get; set; }
            public virtual string ContestYearSpan { get; set; }
        }

        public enum enumReleaseDocType
        {
            Agent,
            School
        }
        #endregion

        #region LABELS
        class labelmodel : Model.Address
        {
            public virtual string Attention { get; set; }
        }

        public class MailLabelPlaceHolder
        {
            public const string Header = "[Header#]";
            public const string Attention = "[Attention#]";
            public const string CustomerName = "[CustomerName#]";
            public const string StreetAddress = "[StreetAddress#]";
            public const string StreetAddressDetail = "[StreetAddressDetail#]";
        }

        public enum enumLabeltype
        {
            allschool,
            allparticipants,
            educators,
            winners,
            judges,
            sponsors,
            firstplaceAgentLabel,
            secondplaceAgentLabel,
            thirdplaceAgentLabel,
            firstplaceSchoolLabel,
            secondplaceSchoolLabel,
            thirdplaceSchoolLabel,
            allschoolwinners
        }
        #endregion

        #region CERTIFICATES
        public class CertificatePlaceHolder
        {
            public const string StudentName = "[studentname]";
            public const string SchoolName = "[schoolname]";
            public const string Placement = "[placement]";
        }

        public class CertificateModel
        {
            public virtual string StudentName { get; set; }
            public virtual string Placement { get; set; }
            public virtual string SchoolName { get; set; }
        }

        public enum CertificateType
        {
            Sponsored,
            NonSponsored
        }
        #endregion

        #region TOP TEN STATE WINNER DOCUMENTS
        public class TopTenDocumentPlaceHolders
        {
            public const string toptenstudentfirstname = "[studentfname]";
            public const string toptenstudentlastname = "[studentlname]";
            public const string placement = "[placement]";
            public const string son_daughter = "[son/daughter]";
            public const string parentnames = "[parentnames]";
            public const string parentcity = "[parentcity]";
            public const string studentaddress = "[studentaddress]";
            public const string studentcity = "[studentcity]";
            public const string studentzip = "[studentzip]";
            public const string teacher = "[teacher]";
            public const string him_her = "[him_her]";
            public const string date = "[date]";
            public const string toptenstudent = "[topten#student]";
            public const string toptenschool = "[topten#school]";
            public const string schoolname = "[schoolname]";
            public const string schooladdress = "[schooladdress]";
            public const string schoolcity = "[schoolcity]";
            public const string schoolzip = "[schoolzip]";
        }
        public class TopTenDocumentModel
        {
            public virtual string toptenstudentfirstname { get; set; }
            public virtual string toptenstudentlastname { get; set; }
            public virtual string placement { get; set; }
            public virtual string son_daughter { get; set; }
            public virtual string parentnames { get; set; }
            public virtual string parentcity { get; set; }
            public virtual string schoolname { get; set; }
            public virtual string schooladdress { get; set; }
            public virtual string schoolzip { get; set; }
            public virtual string schoolcity { get; set; }
            public virtual string contesttopic { get; set; }
            public virtual string contestyearspan { get; set; }
            public virtual string teachername { get; set; }
            public virtual string studentaddress { get; set; }
            public virtual string studentcity { get; set; }
            public virtual string studentzip { get; set; }
            public virtual string him_her { get; set; }
        }
        #endregion
        #endregion
    }
}