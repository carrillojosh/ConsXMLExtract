using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace ConsXMLExtract
{
    /// <summary>
    /// Task
    /// 
    /// Revision History
    /// ----------------
    /// </summary>
    public class ProjectContent
    {
        private string wbs;
        private string name;
        private System.DateTime startDate;
        private System.DateTime finishDate;
        private double percentComplete;
        private string savedVersion;
        private string fileName;
        private string titleName;
        private System.DateTime createdDate;
        private System.DateTime lastSaved;

        public ProjectContent()
        {
            Wbs = "";
            Name = "";
            StartDate = System.DateTime.Now;
            FinishDate = System.DateTime.Now;
            percentComplete = 0.0;
        }

        public ProjectContent(MSProject.Project o)
        {
            SetProperties(o);
        }

        #region Properties
        /// <summary>
        /// Get/Set the work breakdown structure id.
        /// </summary>
        public string Wbs
        {
            get { return wbs; }
            set { wbs = value; }
        }

        /// <summary>
        /// Get/Set the name.
        /// </summary>
        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        /// <summary>
        /// Get/Set the start date.
        /// </summary>
        public System.DateTime StartDate
        {
            get { return startDate; }
            set { startDate = value; }
        }

        /// <summary>
        /// Get the start date in a format that is used by Excel.
        /// </summary>
        public string StartDateExcelFormat
        {
            get
            {
                // the excel format is: 1970-01-01T00:00:00.000
                string s = StartDate.Year.ToString() + "-";
                if (StartDate.Month > 9)
                {
                    s += StartDate.Month.ToString() + "-";
                }
                else
                {
                    s += "0" + StartDate.Month.ToString() + "-";
                }

                if (StartDate.Day > 9)
                {
                    s += StartDate.Day.ToString();
                }
                else
                {
                    s += "0" + StartDate.Day.ToString();
                }
                s += "T00:00:00.000";
                return s;
            }
        }

        /// <summary>
        /// Get/Set the finish date.
        /// </summary>
        public System.DateTime FinishDate
        {
            get { return finishDate; }
            set { finishDate = value; }
        }

        /// <summary>
        /// Get the finish date in a format that is used by Excel.
        /// </summary>
        public string FinishDateExcelFormat
        {
            get
            {
                // the excel format is: 1970-01-01T00:00:00.000
                string s = FinishDate.Year.ToString() + "-";
                if (FinishDate.Month > 9)
                {
                    s += FinishDate.Month.ToString() + "-";
                }
                else
                {
                    s += "0" + FinishDate.Month.ToString() + "-";
                }

                if (FinishDate.Day > 9)
                {
                    s += FinishDate.Day.ToString();
                }
                else
                {
                    s += "0" + FinishDate.Day.ToString();
                }
                s += "T00:00:00.000";
                return s;
            }
        }

        /// <summary>
        /// Get/Set the percent complete.
        /// </summary>
        public double PercentComplete
        {
            get { return percentComplete; }
            set { percentComplete = value; }
        }
        #endregion Properties

        /// <summary>
        /// Assigns all the properties based upon the data in the 
        /// Microsoft Project task.
        /// </summary>
        /// <param name="o">The task to get the values from.</param>
        public void SetProperties(MSProject.Project o)
        {
            Wbs = o.WBS.ToString();
            Name = o.Name.ToString();
            StartDate = (System.DateTime)o.Start;
            FinishDate = (System.DateTime)o.Finish;
            PercentComplete = (System.Int16)o.PercentComplete;
            savedVersion = "1";
            fileName = o.Name;
            titleName = o.Title.ToString();
            createdDate = (System.DateTime)o.CreationDate;
            lastSaved = (System.DateTime)o.LastSaveDate;





        }

        /// <summary>
        /// Returns the task data as XML formatted for Excel.
        /// </summary>
        /// <returns>The task data as XML formatted for Excel.</returns>
        public string ToXml()
        {
            string s = "";

            //s += "<Row>" + System.Environment.NewLine;
            //s += "<Cell><Data ss:Type=\"String\">" + Wbs + "</Data></Cell>" + System.Environment.NewLine;
            //s += "<Cell><Data ss:Type=\"String\">" + ToXml(Name) + "</Data></Cell>" + System.Environment.NewLine;
            //s += "<Cell ss:StyleID=\"s21\"><Data ss:Type=\"DateTime\">" + StartDateExcelFormat + "</Data></Cell>" + System.Environment.NewLine;
            //s += "<Cell ss:StyleID=\"s21\"><Data ss:Type=\"DateTime\">" + FinishDateExcelFormat + "</Data></Cell>" + System.Environment.NewLine;
            //s += "<Cell><Data ss:Type=\"Number\">" + (PercentComplete /100.0) + "</Data></Cell>" + System.Environment.NewLine;
            //s += "</Row>" + System.Environment.NewLine;

            s += "<SaveVersion>" + savedVersion + "</SaveVersion>" + System.Environment.NewLine;
            s += "<Name>" + ToXml(Name) + "</Name>" + System.Environment.NewLine;
            s += "<Title>" + ToXml(titleName) + "</Title>" + System.Environment.NewLine;
            s += "<CreationDate>" + createdDate.ToString("yyyy-MM-dd'T'HH:mm:ss") + "</CreationDate>" + System.Environment.NewLine;
            s += "<LastSaved>" + lastSaved.ToString("yyyy-MM-dd'T'HH:mm:ss") + "</LastSaved>" + System.Environment.NewLine;
            s += "<ScheduleFromStart>1</ScheduleFromStart>" + System.Environment.NewLine;
            s += "<StartDate>" + StartDate.ToString("yyyy-MM-dd'T'HH:mm:ss") + "</StartDate>" + System.Environment.NewLine;
            s += "<FinishDate>" + FinishDate.ToString("yyyy-MM-dd'T'HH:mm:ss") + "</FinishDate>" + System.Environment.NewLine;


            s += "<FYStartDate>1</FYStartDate>" + System.Environment.NewLine;
            s += "<CriticalSlackLimit>0</CriticalSlackLimit>" + System.Environment.NewLine;
            s += "<CurrencyDigits>2</CurrencyDigits>" + System.Environment.NewLine;
            s += "<CurrencySymbol>$</CurrencySymbol>" + System.Environment.NewLine;
            s += "<CurrencyCode>USD</CurrencyCode>" + System.Environment.NewLine;
            s += "<CurrencySymbolPosition>0</CurrencySymbolPosition>" + System.Environment.NewLine;
            s += "<CalendarUID>1</CalendarUID>" + System.Environment.NewLine;
            s += "<DefaultStartTime>08:00:00</DefaultStartTime>" + System.Environment.NewLine;
            s += "<DefaultFinishTime>17:00:00</DefaultFinishTime>" + System.Environment.NewLine;
            s += "<MinutesPerDay>480</MinutesPerDay>" + System.Environment.NewLine;
            s += "<MinutesPerWeek>2400</MinutesPerWeek>" + System.Environment.NewLine;
            s += "<DaysPerMonth>20</DaysPerMonth>" + System.Environment.NewLine;
            s += "<DefaultTaskType>0</DefaultTaskType>" + System.Environment.NewLine;
            s += "<DefaultFixedCostAccrual>3</DefaultFixedCostAccrual>" + System.Environment.NewLine;
            s += "<DefaultStandardRate>0</DefaultStandardRate>" + System.Environment.NewLine;
            s += "<DefaultOvertimeRate>0</DefaultOvertimeRate>" + System.Environment.NewLine;
            s += "<DurationFormat>7</DurationFormat>" + System.Environment.NewLine;
            s += "<WorkFormat>2</WorkFormat>" + System.Environment.NewLine;
            s += "<EditableActualCosts>0</EditableActualCosts>" + System.Environment.NewLine;
            s += "<HonorConstraints>0</HonorConstraints>" + System.Environment.NewLine;
            s += "<InsertedProjectsLikeSummary>1</InsertedProjectsLikeSummary>" + System.Environment.NewLine;
            s += "<MultipleCriticalPaths>0</MultipleCriticalPaths>" + System.Environment.NewLine;
            s += "<NewTasksEffortDriven>0</NewTasksEffortDriven>" + System.Environment.NewLine;
            s += "<NewTasksEstimated>1</NewTasksEstimated>" + System.Environment.NewLine;
            s += "<SplitsInProgressTasks>1</SplitsInProgressTasks>" + System.Environment.NewLine;
            s += "<SpreadActualCost>0</SpreadActualCost>" + System.Environment.NewLine;
            s += "<SpreadPercentComplete>0</SpreadPercentComplete>" + System.Environment.NewLine;
            s += "<TaskUpdatesResource>1</TaskUpdatesResource>" + System.Environment.NewLine;
            s += "<FiscalYearStart>0</FiscalYearStart>" + System.Environment.NewLine;
            s += "<WeekStartDay>0</WeekStartDay>" + System.Environment.NewLine;
            s += "<MoveCompletedEndsBack>0</MoveCompletedEndsBack>" + System.Environment.NewLine;
            s += "<MoveRemainingStartsBack>0</MoveRemainingStartsBack>" + System.Environment.NewLine;
            s += "<MoveRemainingStartsForward>0</MoveRemainingStartsForward>" + System.Environment.NewLine;
            s += "<MoveCompletedEndsForward>0</MoveCompletedEndsForward>" + System.Environment.NewLine;
            s += "<BaselineForEarnedValue>0</BaselineForEarnedValue>" + System.Environment.NewLine;
            s += "<AutoAddNewResourcesAndTasks>1</AutoAddNewResourcesAndTasks>" + System.Environment.NewLine;
            s += "<CurrentDate>" + DateTime.Now.ToString("yyyy-MM-dd'T'HH:mm:ss") + "</CurrentDate>" + System.Environment.NewLine;
            s += "<MicrosoftProjectServerURL>1</MicrosoftProjectServerURL>" + System.Environment.NewLine;
            s += "<Autolink>0</Autolink>" + System.Environment.NewLine;
            s += "<NewTaskStartDate>0</NewTaskStartDate>" + System.Environment.NewLine;
            s += "<NewTasksAreManual>1</NewTasksAreManual>" + System.Environment.NewLine;
            s += "<DefaultTaskEVMethod>0</DefaultTaskEVMethod>" + System.Environment.NewLine;
            s += "<ProjectExternallyEdited>0</ProjectExternallyEdited>" + System.Environment.NewLine;
            s += "<ExtendedCreationDate>1984-01-01T00:00:00</ExtendedCreationDate>" + System.Environment.NewLine;
            s += "<ActualsInSync>0</ActualsInSync>" + System.Environment.NewLine;
            s += "<RemoveFileProperties>0</RemoveFileProperties>" + System.Environment.NewLine;
            s += "<AdminProject>0</AdminProject>" + System.Environment.NewLine;
            s += "<UpdateManuallyScheduledTasksWhenEditingLinks>1</UpdateManuallyScheduledTasksWhenEditingLinks>" + System.Environment.NewLine;
            s += "<KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled>0</KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled>" + System.Environment.NewLine;
            s += "<OutlineCodes/>" + System.Environment.NewLine;
            s += "<WBSMasks/>" + System.Environment.NewLine;

            s += "<ExtendedAttributes>" + System.Environment.NewLine;
            s += "<ExtendedAttribute>" + System.Environment.NewLine;
            s += "<FieldID>188743731</FieldID>" + System.Environment.NewLine;
            s += "<FieldName>Text1</FieldName>" + System.Environment.NewLine;
            s += "<Alias>Required Artifact</Alias>" + System.Environment.NewLine;
            s += "<Guid>000039B7-8BBE-4CEB-82C4-FA8C0B400033</Guid>" + System.Environment.NewLine;
            s += "<SecondaryPID>255869028</SecondaryPID>" + System.Environment.NewLine;
            s += "<SecondaryGuid>000039B7-8BBE-4CEB-82C4-FA8C0F404064</SecondaryGuid>" + System.Environment.NewLine;
            s += "</ExtendedAttribute>" + System.Environment.NewLine;
            s += "</ExtendedAttributes>" + System.Environment.NewLine;

            s += "<Calendars>" + System.Environment.NewLine;
            s += "<Calendar>" + System.Environment.NewLine;
            s += "<UID>1</UID>" + System.Environment.NewLine;
            s += "<Name>Standard</Name>" + System.Environment.NewLine;
            s += "<IsBaseCalendar>1</IsBaseCalendar>" + System.Environment.NewLine;
            s += "<IsBaselineCalendar>0</IsBaselineCalendar>" + System.Environment.NewLine;
            s += "<BaseCalendarUID>-1</BaseCalendarUID>" + System.Environment.NewLine;
            s += "<WeekDays>" + System.Environment.NewLine;
            s += "<WeekDay>" + System.Environment.NewLine;
            s += "<DayType>1</DayType>" + System.Environment.NewLine;
            s += "<DayWorking>0</DayWorking>" + System.Environment.NewLine;
            s += "</WeekDay>" + System.Environment.NewLine;
            s += "<WeekDay>" + System.Environment.NewLine;
            s += "<DayType>2</DayType>" + System.Environment.NewLine;
            s += "<DayWorking>1</DayWorking>" + System.Environment.NewLine;
            s += "<WorkingTimes>" + System.Environment.NewLine;
            s += "<WorkingTime>" + System.Environment.NewLine;
            s += "<FromTime>08:00:00</FromTime>" + System.Environment.NewLine;
            s += "<ToTime>12:00:00</ToTime>" + System.Environment.NewLine;
            s += "</WorkingTime>" + System.Environment.NewLine;
            s += "<WorkingTime>" + System.Environment.NewLine;
            s += "<FromTime>13:00:00</FromTime>" + System.Environment.NewLine;
            s += "<ToTime>17:00:00</ToTime>" + System.Environment.NewLine;
            s += "</WorkingTime>" + System.Environment.NewLine;
            s += "</WorkingTimes>" + System.Environment.NewLine;
            s += "</WeekDay>" + System.Environment.NewLine;
            s += "</WeekDays>" + System.Environment.NewLine;
            s += "<Exceptions>" + System.Environment.NewLine;
            s += "<Exception>" + System.Environment.NewLine;
            s += "<EnteredByOccurrences>0</EnteredByOccurrences>" + System.Environment.NewLine;
            s += "<TimePeriod>" + System.Environment.NewLine;
            s += "<FromDate>2017-11-18T00:00:00</FromDate>" + System.Environment.NewLine;
            s += "<ToDate>2017-11-18T23:59:00</ToDate>" + System.Environment.NewLine;
            s += "</TimePeriod>" + System.Environment.NewLine;
            s += "<Occurrences>1</Occurrences>" + System.Environment.NewLine;
            s += "<Name></Name>" + System.Environment.NewLine;
            s += "<Type>1</Type>" + System.Environment.NewLine;
            s += "<DayWorking>1</DayWorking>" + System.Environment.NewLine;
            s += "<WorkingTimes>" + System.Environment.NewLine;
            s += "<WorkingTime>" + System.Environment.NewLine;
            s += "<FromTime>08:00:00</FromTime>" + System.Environment.NewLine;
            s += "<ToTime>12:00:00</ToTime>" + System.Environment.NewLine;
            s += "</WorkingTime>" + System.Environment.NewLine;
            s += "<WorkingTime>" + System.Environment.NewLine;
            s += "<FromTime>13:00:00</FromTime>" + System.Environment.NewLine;
            s += "<ToTime>17:00:00</ToTime>" + System.Environment.NewLine;
            s += "</WorkingTime>" + System.Environment.NewLine;
            s += "</WorkingTimes>" + System.Environment.NewLine;
            s += "</Exception>" + System.Environment.NewLine;
            s += "</Exceptions>" + System.Environment.NewLine;
            s += "</Calendar>" + System.Environment.NewLine;
            s += "<Calendar>" + System.Environment.NewLine;
            s += "<UID>3</UID>" + System.Environment.NewLine;
            s += "<Name>Yan (Yana</Name>" + System.Environment.NewLine;
            s += "<IsBaseCalendar>0</IsBaseCalendar>" + System.Environment.NewLine;
            s += "<IsBaselineCalendar>0</IsBaselineCalendar>" + System.Environment.NewLine;
            s += "<BaseCalendarUID>1</BaseCalendarUID>" + System.Environment.NewLine;
            s += "</Calendar>" + System.Environment.NewLine;
            s += "</Calendars>";


            // *** always remove last newline reference


            return s;
        }

        /// <summary>
        /// Converts the passed string into a string that can be used in XML.  The conversions are:
        /// the ampersand, less than, greater than quote, and appostrophe.
        /// </summary>
        /// <param name="str">The string to convert.</param>
        /// <returns>The converted string.</returns>
        public static string ToXml(string str)
        {
            string s = (str == null) ? "" : str;
            s = s.Trim().Replace("&", "&#38;");
            s = s.Replace("?", "&#63;");
            s = s.Replace("<", "&lt;");
            s = s.Replace(">", "&gt;");
            s = s.Replace("\"", "&quot;");
            s = s.Replace("'", "&apos;");
            s = s.Replace("~", "&#126;");
            return s;
        }
    }
}
