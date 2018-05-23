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
    public class Task
    {
        private string wbs;
        private string name;
        private System.DateTime startDate;
        private System.DateTime finishDate;
        private double percentComplete;
        private int uid;
        private int id;
        private bool active;
        private System.DateTime createdDate;
        private string outlineNumber;
        private int outlineLevel;
        private DateTime start;
        private DateTime finish;



        public Task()
        {

            Name = "";
            StartDate = System.DateTime.Now;
            FinishDate = System.DateTime.Now;
            PercentComplete = 0.0;
            UID = 0;
            ID = 0;
            Active = true;
            CreatedDate = System.DateTime.Now;
            Wbs = "";
            OutlineNumber = "";
            OutlineLevel = 0;
            Start = System.DateTime.Now;
            Finish = System.DateTime.Now;
        }

        public Task(MSProject.Task o)
        {
            SetProperties(o);
        }

        // *****************************
        #region Properties Values
        /// <summary>
        /// Get/Set the work breakdown structure id.
        /// </summary>
        public int OutlineLevel
        {
            get { return outlineLevel; }
            set { outlineLevel = value; }
        }
        public DateTime Start
        {
            get { return start; }
            set { start = value; }
        }
        public DateTime Finish
        {
            get { return finish; }
            set { finish = value; }
        }
        public string Wbs
        {
            get { return wbs; }
            set { wbs = value; }
        }

        public int UID
        {
            get { return uid; }
            set { uid = value; }
        }
        public int ID
        {
            get { return id; }
            set { id = value; }
        }

        public bool Active
        {
            get { return active; }
            set { active = value; }
        }
        public System.DateTime CreatedDate
        {
            get { return createdDate; }
            set { createdDate = value; }
        }
        public string OutlineNumber
        {
            get { return outlineNumber; }
            set { outlineNumber = value; }
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
        #endregion Properties Values

        /// <summary>
        /// Assigns all the properties based upon the data in the 
        /// Microsoft Project task.
        /// </summary>
        /// <param name="o">The task to get the values from.</param>
        public void SetProperties(MSProject.Task o)
        {

            Name = o.Name.ToString();
            PercentComplete = (System.Int16)o.PercentComplete;
            UID = (System.Int16)o.UniqueID;
            ID = (System.Int16)o.ID;
            Active = (System.Boolean)o.Active;
            if (o.Created.ToString() != "NA")
            {
                CreatedDate = (System.DateTime)o.Created;
            }
            Wbs = o.WBS.ToString();
            OutlineNumber = o.OutlineNumber;
            OutlineLevel = o.OutlineLevel;
            if (o.Start.ToString() != "NA")
            {
                Start = (DateTime)o.Start;
            }
            if (o.Finish.ToString() != "NA")
            {
                Finish = (DateTime)o.Finish;
            }

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

            s += "<Task>" + System.Environment.NewLine;
            s += "<UID>" + UID + "</UID>" + System.Environment.NewLine;
            s += "<ID>" + ID + "</ID>" + System.Environment.NewLine;
            s += "<Name>" + ToXml(Name) + "</Name>" + System.Environment.NewLine;
            s += "<Active>" + Convert.ToInt16(Active) + "</Active>" + System.Environment.NewLine;
            s += "<Manual>0</Manual>" + System.Environment.NewLine;
            s += "<Type>0</Type>" + System.Environment.NewLine;
            s += "<IsNull>0</IsNull>" + System.Environment.NewLine;
            s += "<CreateDate>" + CreatedDate.ToString("yyyy-MM-dd'T'HH:mm:ss") + "</CreateDate>" + System.Environment.NewLine;
            s += "<WBS>" + Wbs + "</WBS>" + System.Environment.NewLine;
            s += "<OutlineNumber>" + OutlineNumber + "</OutlineNumber>" + System.Environment.NewLine;
            s += "<OutlineLevel>" + OutlineLevel + "</OutlineLevel>" + System.Environment.NewLine;

            s += "<Priority>500</Priority>" + System.Environment.NewLine;

            s += "<Start>" + Start.ToString("yyyy-MM-dd'T'HH:mm:ss") + "</Start>" + System.Environment.NewLine;
            s += "<Finish>" + Finish.ToString("yyyy-MM-dd'T'HH:mm:ss") + "</Finish>" + System.Environment.NewLine;

            s += "<Duration>PT64H0M0S</Duration>" + System.Environment.NewLine;
            s += "<ManualStart>2017-12-01T08:00:00</ManualStart>" + System.Environment.NewLine;
            s += "<ManualFinish>2017-12-12T17:00:00</ManualFinish>" + System.Environment.NewLine;
            s += "<ManualDuration>PT64H0M0S</ManualDuration>" + System.Environment.NewLine;
            s += "<DurationFormat>7</DurationFormat>" + System.Environment.NewLine;
            s += "<Work>PT0H0M0S</Work>" + System.Environment.NewLine;
            s += "<Stop>2017-12-12T17:00:00</Stop>" + System.Environment.NewLine;
            s += "<Resume>2017-12-12T17:00:00</Resume>" + System.Environment.NewLine;
            s += "<ResumeValid>0</ResumeValid>" + System.Environment.NewLine;
            s += "<EffortDriven>0</EffortDriven>" + System.Environment.NewLine;
            s += "<Recurring>0</Recurring>" + System.Environment.NewLine;
            s += "<OverAllocated>0</OverAllocated>" + System.Environment.NewLine;
            s += "<Estimated>0</Estimated>" + System.Environment.NewLine;
            s += "<Milestone>0</Milestone>" + System.Environment.NewLine;
            s += "<Summary>0</Summary>" + System.Environment.NewLine;
            s += "<DisplayAsSummary>0</DisplayAsSummary>" + System.Environment.NewLine;
            s += "<Critical>0</Critical>" + System.Environment.NewLine;
            s += "<IsSubproject>0</IsSubproject>" + System.Environment.NewLine;
            s += "<IsSubprojectReadOnly>0</IsSubprojectReadOnly>" + System.Environment.NewLine;
            s += "<ExternalTask>0</ExternalTask>" + System.Environment.NewLine;
            s += "<EarlyStart>2017-12-01T08:00:00</EarlyStart>" + System.Environment.NewLine;
            s += "<EarlyFinish>2017-12-12T17:00:00</EarlyFinish>" + System.Environment.NewLine;
            s += "<LateStart>2017-12-01T08:00:00</LateStart>" + System.Environment.NewLine;
            s += "<LateFinish>2017-12-12T17:00:00</LateFinish>" + System.Environment.NewLine;
            s += "<StartVariance>0</StartVariance>" + System.Environment.NewLine;
            s += "<FinishVariance>0</FinishVariance>" + System.Environment.NewLine;
            s += "<WorkVariance>0.00</WorkVariance>" + System.Environment.NewLine;
            s += "<FreeSlack>0</FreeSlack>" + System.Environment.NewLine;
            s += "<TotalSlack>0</TotalSlack>" + System.Environment.NewLine;
            s += "<StartSlack>0</StartSlack>" + System.Environment.NewLine;
            s += "<FinishSlack>0</FinishSlack>" + System.Environment.NewLine;
            s += "<FixedCost>0</FixedCost>" + System.Environment.NewLine;
            s += "<FixedCostAccrual>3</FixedCostAccrual>" + System.Environment.NewLine;

            s += "<PercentComplete>" + PercentComplete + "</PercentComplete>" + System.Environment.NewLine;

            s += "<PercentWorkComplete>100</PercentWorkComplete>" + System.Environment.NewLine;
            s += "<Cost>0</Cost>" + System.Environment.NewLine;
            s += "<OvertimeCost>0</OvertimeCost>" + System.Environment.NewLine;
            s += "<OvertimeWork>PT0H0M0S</OvertimeWork>" + System.Environment.NewLine;
            s += "<ActualStart>2017-12-01T08:00:00</ActualStart>" + System.Environment.NewLine;
            s += "<ActualFinish>2017-12-12T17:00:00</ActualFinish>" + System.Environment.NewLine;
            s += "<ActualDuration>PT64H0M0S</ActualDuration>" + System.Environment.NewLine;
            s += "<ActualCost>0</ActualCost>" + System.Environment.NewLine;
            s += "<ActualOvertimeCost>0</ActualOvertimeCost>" + System.Environment.NewLine;
            s += "<ActualWork>PT0H0M0S</ActualWork>" + System.Environment.NewLine;
            s += "<ActualOvertimeWork>PT0H0M0S</ActualOvertimeWork>" + System.Environment.NewLine;
            s += "<RegularWork>PT0H0M0S</RegularWork>" + System.Environment.NewLine;
            s += "<RemainingDuration>PT0H0M0S</RemainingDuration>" + System.Environment.NewLine;
            s += "<RemainingCost>0</RemainingCost>" + System.Environment.NewLine;
            s += "<RemainingWork>PT0H0M0S</RemainingWork>" + System.Environment.NewLine;
            s += "<RemainingOvertimeCost>0</RemainingOvertimeCost>" + System.Environment.NewLine;
            s += "<RemainingOvertimeWork>PT0H0M0S</RemainingOvertimeWork>" + System.Environment.NewLine;
            s += "<ACWP>0.00</ACWP>" + System.Environment.NewLine;
            s += "<CV>0.00</CV>" + System.Environment.NewLine;
            s += "<ConstraintType>6</ConstraintType>" + System.Environment.NewLine;
            s += "<CalendarUID>-1</CalendarUID>" + System.Environment.NewLine;
            s += "<ConstraintDate>2017-12-12T17:00:00</ConstraintDate>" + System.Environment.NewLine;
            s += "<LevelAssignments>1</LevelAssignments>" + System.Environment.NewLine;
            s += "<LevelingCanSplit>1</LevelingCanSplit>" + System.Environment.NewLine;
            s += "<LevelingDelay>0</LevelingDelay>" + System.Environment.NewLine;
            s += "<LevelingDelayFormat>8</LevelingDelayFormat>" + System.Environment.NewLine;
            s += "<IgnoreResourceCalendar>0</IgnoreResourceCalendar>" + System.Environment.NewLine;
            s += "<HideBar>0</HideBar>" + System.Environment.NewLine;
            s += "<Rollup>0</Rollup>" + System.Environment.NewLine;
            s += "<BCWS>0.00</BCWS>" + System.Environment.NewLine;
            s += "<BCWP>0.00</BCWP>" + System.Environment.NewLine;
            s += "<PhysicalPercentComplete>0</PhysicalPercentComplete>" + System.Environment.NewLine;
            s += "<EarnedValueMethod>0</EarnedValueMethod>" + System.Environment.NewLine;
            s += "<PredecessorLink>" + System.Environment.NewLine;
            s += "  <PredecessorUID>147</PredecessorUID>" + System.Environment.NewLine;
            s += "  <Type>1</Type>" + System.Environment.NewLine;
            s += "  <CrossProject>0</CrossProject>" + System.Environment.NewLine;
            s += "  <LinkLag>0</LinkLag>" + System.Environment.NewLine;
            s += "  <LagFormat>7</LagFormat>" + System.Environment.NewLine;
            s += "</PredecessorLink>" + System.Environment.NewLine;
            s += "<IsPublished>1</IsPublished>" + System.Environment.NewLine;
            s += "<CommitmentType>0</CommitmentType>" + System.Environment.NewLine;
            s += "<ExtendedAttribute>" + System.Environment.NewLine;
            s += "  <FieldID>188743731</FieldID>" + System.Environment.NewLine;
            s += "  <Value>Yes</Value>" + System.Environment.NewLine;
            s += "</ExtendedAttribute>" + System.Environment.NewLine;
            s += "  <TimephasedData>" + System.Environment.NewLine;
            s += "      <Type>11</Type>" + System.Environment.NewLine;
            s += "      <UID>149</UID>" + System.Environment.NewLine;
            s += "      <Start>2017-12-12T08:00:00</Start>" + System.Environment.NewLine;
            s += "      <Finish>2017-12-12T17:00:00</Finish>" + System.Environment.NewLine;
            s += "      <Unit>2</Unit>" + System.Environment.NewLine;
            s += "      <Value>12.5</Value>" + System.Environment.NewLine;
            s += "  </TimephasedData>" + System.Environment.NewLine;
            s += "</Task>";





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
