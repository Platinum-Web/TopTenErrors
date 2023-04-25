using System;
namespace TopTenErrors.Models
{
	public class ExcelFile
	{
		public string REPORT_DATE { get; set; }
		public string PROGRAM { get; set; }
		public string WEEK_ENDING { get; set; }
		public string ORIG_CTR { get; set; }
		public string ERROR_NUMBER { get; set; }
		public string NO_OF_ERRORS { get; set; }
		public string ERROR_MESSAGE { get; set; }
		public string MONTH_YEAR { get; set; }

        public ExcelFile(string rEPORT_DATE, string pROGRAM, string wEEK_ENDING, string oRIG_CTR, string eRROR_NUMBER, string nO_OF_ERRORS, string nO_MESSAGE, string mONTH_YEAR)
        {
            REPORT_DATE = rEPORT_DATE;
            PROGRAM = pROGRAM;
            WEEK_ENDING = wEEK_ENDING;
            ORIG_CTR = oRIG_CTR;
            ERROR_NUMBER = eRROR_NUMBER;
            NO_OF_ERRORS = nO_OF_ERRORS;
            ERROR_MESSAGE = nO_MESSAGE;
            MONTH_YEAR = mONTH_YEAR;
        }
    }
}

