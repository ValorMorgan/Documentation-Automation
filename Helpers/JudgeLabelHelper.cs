using FB.AmericaMe.UI.Helpers.Documents;
using FB.Common.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FB.AmericaMe.UI.Helpers
{
    public class JudgeLabelHelper
    {
        public static IList<JudgeLabelModel> getJudgeLabelModels(IList<string> jdgs)
        {
            IList<JudgeLabelModel> jdglblmodels = new List<JudgeLabelModel>();

            var allparticipantjudges = Helpers.MainWindowHelper.MainWindow.ParticipantJudgeController.ParticipantJudges;
            var date = DateTime.Now.Date.ToString("MM/dd/yyyy");

            foreach (var jdg in jdgs)
            {
                var judgeparticipants = allparticipantjudges.Where(p => p.Judge.Person.FullName.Equals(jdg))
                    .Select(s => s.Participant).ToList();


                foreach (var participant in judgeparticipants)
                {
                    jdglblmodels.Add(new JudgeLabelModel()
                        {
                            JudgeName = string.Format("Materials for {0}",jdg),
                            SchoolName = participant.School.Name,
                            SchoolCity = participant.School.Address.City,
                            SchoolAddress = participant.School.Address.Address1 + " " + participant.School.Address.Address2,
                            SchoolZip = participant.School.Address.Zip,
                            NumOfEssays = string.Format("Essays Submitted: {0}", participant.EssayCount),
                            Date = date
                        });
                }
            }

            return jdglblmodels;
        }

        public class JudgeLabelModel
        {
            public virtual string JudgeName { get; set; }
            public virtual string SchoolName { get; set; }
            public virtual string SchoolCity { get; set; }
            public virtual string SchoolAddress { get; set; }
            public virtual string SchoolZip { get; set; }
            public virtual string NumOfEssays { get; set; }
            public virtual string Date { get; set; }
        }

        public class JudgeLabelPlaceHolder
        {
            public const string Judge = "[judge#]";
            public const string SchoolName = "[schoolname#]";
            public const string SchoolCity = "[schoolcity#]";
            public const string NumberOfEssays = "[numofessays#]";
            public const string Date = "[date#]";
        }
    }
}
