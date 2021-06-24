using System;
using System.Collections.Generic;
using System.Text;

namespace EnrollmentExcelTool
{
    public class ProfileFileInfo
    {
        public bool isIdColumnExists { get; set; }
        public bool isCampCodeColumnExists { get; set; }
        public bool isStatusColumnExists { get; set; }
        public bool isPnOnColumnExists { get; set; }        
        public bool isSnOnColumnExists { get; set; }
        public bool isWorkPerformedICAOColumnExists { get; set; }
        public bool isPageColumnExists { get; set; }
        public bool isValid { get; set; }
    }
}
