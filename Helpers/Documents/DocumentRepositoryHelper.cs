using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace FB.AmericaMe.UI.Helpers.Documents
{

    public class DocumentTemplates
    {
        public const string AgentConfirmationLetter = "AgentConfirmationLetter";
        
        public const string AgentMemo1 = "AgentMemo1";
        public const string AgentMemo2 = "AgentMemo2";
        public const string AgentMemo3 = "AgentMemo3";
        
        public const string AgentNewsRelease1 = "AgentNewsRelease1";
        public const string AgentNewsRelease2 = "AgentNewsRelease2";
        public const string AgentNewsRelease3 = "AgentNewsRelease3";
        public const string AgentNewsRelease123 = "AgentNewsRelease123";

        public const string AgentTeacherLetter1 = "AgentTeachLetter1";
        public const string AgentTeacherLetter2 = "AgentTeachLetter2";
        public const string AgentTeacherLetter3 = "AgentTeachLetter3";
        public const string AgentTeacherLetter123 = "AgentTeachLetter123";
        
        public const string SchoolNewsRelease1 = "SchoolNewsRelease1";
        public const string SchoolNewsRelease2 = "SchoolNewsRelease2";
        public const string SchoolNewsRelease3 = "SchoolNewsRelease3";
        public const string SchoolNewsRelease123 = "SchoolNewsRelease123";

        public const string SchoolTeacherLetter1 = "SchoolTeachLetter1";
        public const string SchoolTeacherLetter2 = "SchoolTeachLetter2";
        public const string SchoolTeacherLetter3 = "SchoolTeachLetter3";
        public const string SchoolTeacherLetter123 = "SchoolTeachLetter123";
                      
        public const string Certificate = "Certificate";
        
        public const string Labels = "Labels";
        public const string JudgeLabels = "JudgeLabels";

        public const string TeacherConfirmationLetter = "TeacherConfirmationLetter";

        public const string TopTenNewsRelease   = "TopTenNewsRelease";
        public const string TopTenStudentLetter = "TopTenStudentLetter";
        public const string TopTenTeacherLetter = "TopTenTeacherLetter";

        
    }

    public class DocumentRepositoryHelper
    {
        private const string savedocext = "doc";
        private static string[] saveDocExt = { "doc", "txt", "xlsx" };
        private const string doctemplateext = "rtf";

        public static string WorkingDirectory = new FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName;
        public static string DefaultDirectory = new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)).DirectoryName;
        public static string DocumentTemplateDirectory = string.Format(@"{0}\{1}", WorkingDirectory, "DocumentTemplates");
        public static string SavedTempDocumentsDirectory = string.Format(@"{0}\{1}", WorkingDirectory, "Documents");
        public static string SavedDocumentsDirectory = string.Format(@"{0}\{1}", DefaultDirectory, "Documents");

        public static string GetDocumentTemplate(string docTemp)
        {
            return string.Format(@"{0}\{1}.{2}", DocumentTemplateDirectory, docTemp, doctemplateext);
        }

        #region DocumentFileName
        /// <summary>
        /// Returns a string of the full directory of "SavedTempDocumentsDirectory"
        /// and the provided document name + current date.  The file defaults to
        /// a word document extension.
        /// </summary>
        /// <param name="doc"></param>
        /// <returns>"SavedTempDocumentsDirectory"\"doc".doc</returns>
        public static string TempDocumentFileName(string doc)
        {
            if (!System.IO.Directory.Exists(SavedTempDocumentsDirectory))
                System.IO.Directory.CreateDirectory(SavedTempDocumentsDirectory);
            string rtn = string.Format(@"{0}\{1}.{2}", SavedTempDocumentsDirectory, string.Format("{0}_{1}_{2}", ContestEntryHelper.Contest.StartDate.Year.ToString(), doc, DateTime.Now.ToString("MMddyyyy_HHmm")), saveDocExt[0]);

            return rtn;
        }

        /// <summary>
        /// Returns a string of the full directory of "SavedTempDocumentsDirectory"
        /// and the provided document name + current date.  ext must be "doc", or
        /// "txt" otherwise returns a null.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="ext"></param>
        /// <returns>"SavedTempDocumentsDirectory"\"doc"."ext"</returns>
        public static string TempDocumentFileName(string doc, string ext)
        {
            if (!saveDocExt.Contains(ext))
            {
                Console.WriteLine("Error in FB.AmericaMe.UI.Documents.DocumentRepositoryHelper.cs on TempDocumentFileName");
                Console.WriteLine("File extension not recognized!  Expecting doc, txt, or xlsx.");
                return null;
            }
            if (!System.IO.Directory.Exists(SavedTempDocumentsDirectory))
                System.IO.Directory.CreateDirectory(SavedTempDocumentsDirectory);
            string rtn = string.Format(@"{0}\{1}.{2}", SavedTempDocumentsDirectory, string.Format("{0}_{1}_{2}", ContestEntryHelper.Contest.StartDate.Year.ToString(), doc, DateTime.Now.ToString("MMddyyyy_HHmm")), ext);

            return rtn;
        }

        /// <summary>
        /// Returns a string of the full directory of "SavedDocumentsDirectory"
        /// and the provided document name + current date.  Uses a default extension
        /// of doc.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="ext"></param>
        /// <returns>"SavedDocumentsDirectory"\"doc".doc</returns>
        public static string DocumentFileName(string doc)
        {
            if (!System.IO.Directory.Exists(SavedDocumentsDirectory))
                System.IO.Directory.CreateDirectory(SavedDocumentsDirectory);
            string rtn = string.Format(@"{0}\{1}.{2}", SavedDocumentsDirectory, string.Format("{0}_{1}_{2}", ContestEntryHelper.Contest.StartDate.Year.ToString(), doc, DateTime.Now.ToString("MMddyyyy_HHmm")), saveDocExt[0]);

            return rtn;
        }

        /// <summary>
        /// Returns a string of the full directory of "SavedDocumentsDirectory"
        /// and the provided document name + current date.  ext must be "doc" or
        /// "txt" otherwise returns a null.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="ext"></param>
        /// <returns>"SavedDocumentsDirectory"\"doc"."ext"</returns>
        public static string DocumentFileName(string doc, string ext)
        {
            if (!saveDocExt.Contains(ext))
            {
                Console.WriteLine("Error in FB.AmericaMe.UI.Documents.DocumentRepositoryHelper.cs on DocumentFileName");
                Console.WriteLine("File extension not recognized!  Expecting doc, txt, or xlsx.");
                return null;
            }
            if (!System.IO.Directory.Exists(SavedDocumentsDirectory))
                System.IO.Directory.CreateDirectory(SavedDocumentsDirectory);
            string rtn = string.Format(@"{0}\{1}.{2}", SavedDocumentsDirectory, string.Format("{0}_{1}_{2}", ContestEntryHelper.Contest.StartDate.Year.ToString(), doc, DateTime.Now.ToString("MMddyyyy_HHmm")), ext);

            return rtn;
        }
        #endregion

        public static IList<Model.DocumentTemplate> GetDocumentModels()
        {
            IList<Model.DocumentTemplate> rtnList = new List<Model.DocumentTemplate>();
            rtnList.Add(new Model.DocumentTemplate()
            {
                DocumentName = ""
            });

            var arr = System.IO.Directory.GetFiles(DocumentTemplateDirectory)
                 .Select(path => Path.GetFileNameWithoutExtension(path))
                                     .ToArray();
            foreach (var a in arr)
            {
                rtnList.Add(new Model.DocumentTemplate()
                {
                  DocumentName = a
                });
            }

            return rtnList;
        }

        /// <summary>
        /// Deletes all files in the temporary location that are older
        /// than a week.
        /// </summary>
        public static void DeleteOldFiles()
        {
            if (!System.IO.Directory.Exists(SavedTempDocumentsDirectory))
                return;

            DateTime lastWeek = DateTime.Now.Subtract(TimeSpan.FromDays(7));
            string[] files = System.IO.Directory.GetFiles(SavedTempDocumentsDirectory);

            try
            {
                foreach (string file in files)
                {
                    DateTime fileCreation = System.IO.File.GetCreationTime(file);
                    if (fileCreation < lastWeek)
                    {
                        System.IO.File.Delete(file);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Gets the latest created file in the temporary directory.
        /// If no file exists, returns null;
        /// </summary>
        /// <param name="template">The template name to look for.</param>
        /// <param name="fileNameAddition">Any added section to the template filename (ie _Spreadsheet).</param>
        /// <returns>The full directory and filename of the file.</returns>
        public static string GetLatestFile(string template, string fileNameAddition = null)
        {
            if (!System.IO.Directory.Exists(SavedTempDocumentsDirectory))
                return null;

            string searchedFile = string.Format("{0}_{1}{2}", ContestEntryHelper.Contest.StartDate.Year.ToString(), template, fileNameAddition);
            string[] files = System.IO.Directory.GetFiles(SavedTempDocumentsDirectory);
            string latestFile = null;
            DateTime latestFileCreation = new DateTime();

            try
            {
                foreach (string file in files)
                {
                    if (file.Contains(searchedFile))
                    {
                        DateTime fileCreation = System.IO.File.GetCreationTime(file);
                        if (fileCreation > latestFileCreation)
                        {
                            latestFile = file;
                            latestFileCreation = fileCreation;
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }

            return latestFile;
        }
    }
}
