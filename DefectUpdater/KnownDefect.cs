using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DefectUpdater {
    public class KnownDefect : IEquatable<KnownDefect> {
        public string Project { get; set; }
        public string Upgrade { get; set; }
        public string DefectNo { get; set; }
        public string MasterTransNo { get; set; }
        public string TestTransNo { get; set; }
        public string SecId { get; set; }
        public string DeviationColumnName { get; set; }
        public string MasterValue { get; set; }
        public string TestValue { get; set; }

        public KnownDefect(string project, string upgrade, string defectNo, string masterTransNo, string testTransNo, string secId, string deviationColumnName, string masterValue, string testValue) {
            Project = project.Trim();
            Upgrade = upgrade.Trim();
            DefectNo = defectNo.Trim();
            MasterTransNo = masterTransNo.Trim();
            TestTransNo = testTransNo.Trim();
            SecId = secId.Trim();
            DeviationColumnName = deviationColumnName.Trim();
            MasterValue = masterValue.Trim();
            TestValue = testValue.Trim();
        }

        public bool Equals(KnownDefect other) {
            if (Project == other.Project &&
            Upgrade == other.Upgrade &&
            DefectNo == other.DefectNo &&
            MasterTransNo == other.MasterTransNo &&
            TestTransNo == other.TestTransNo &&
            SecId == other.SecId &&
            DeviationColumnName == other.DeviationColumnName &&
            MasterValue == other.MasterValue &&
            TestValue == other.TestValue) {
                return true;
            } else {
                return false;
            }
        }
    }
}
