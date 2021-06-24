using System;
using System.Collections.Generic;
using System.Text;

namespace EnrollmentExcelTool
{
    public class SourceFileInfo
    {
        public bool isIdColumnExists { get; set; }
        public bool isCompCityInstldColumnExists { get; set; }
        public bool isEnodingConditionColumnExists { get; set; }
        public bool isCampCodeColumnExists { get; set; }
        public bool isDuplicateColumnExists { get; set; }
        public bool isAddressedColumnExists { get; set; }
        public bool isPositionColumnExists { get; set; }
        public bool isDescriptionColumnExists { get; set; }
        public bool isPartNumberColumnExists { get; set; }
        public bool isSerialNumberColumnExists { get; set; }
        public bool isPageColumnExists { get; set; }
        public bool isValid { get; set; }

    }
}
