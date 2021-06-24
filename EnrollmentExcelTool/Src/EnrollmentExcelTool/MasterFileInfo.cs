using System;
using System.Collections.Generic;
using System.Text;

namespace EnrollmentExcelTool
{
    public class MasterFileInfo
    {
        public bool isIdColumnExists { get; set; }
        public bool isCampCodeColumnExists { get; set; }
        public bool isEnodedColumnExists { get; set; }
        public bool isPositionColumnExists { get; set; }
        public bool isDescriptionColumnExists { get; set; }
        public bool isPartNumberColumnExists { get; set; }
        public bool isValid { get; set; }
    }
}
