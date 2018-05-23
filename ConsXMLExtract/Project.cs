using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using MSProject = Microsoft.Office.Interop.MSProject;
using System.IO;

namespace ConsXMLExtract
{
    class Project
    {
        private ArrayList projects;
        private ArrayList tasks;
        private ArrayList current;
        private ArrayList currentProjects;
        private ArrayList future;
        private ArrayList late;
        public String newFileName;


        static void Main(string[] args)
        {

            Project project = new Project();
            string err = project.Load(args[0]);
            if (err.Length > 0)
            {
                Console.Write(err);
            }

            project.EvaluateTasks();

            err = project.Save(args[1]);
            if (err.Length > 0)
            {
                Console.Write(err);
            }
            else
            {
                Console.WriteLine("------------");
                Console.WriteLine("Successful XML File Conversion!");
                Console.WriteLine("------------");
            }
        }




        public string NewFileName
        {
            get { return newFileName; }
            set { newFileName = value; }
        }







        public Project()
        {
            Initialize();
        }

        private void Initialize()
        {
            projects = new ArrayList();
            tasks = new ArrayList();
            current = new ArrayList();
            future = new ArrayList();
            late = new ArrayList();
            currentProjects = new ArrayList();
            newFileName = "";
        }

        /// <summary>
        /// Loads all the tasks from the Microsoft Project file into
        /// this class.
        /// </summary>
        /// <param name="fileName">The full path name of the Microsoft Project file.</param>
        /// <returns>On succes an empty string.  On error, the error description.</returns>
        public string Load(string fileName)
        {
            MSProject.ApplicationClass app = null;
            string retVal = "";
            Initialize();

            newFileName = Path.GetFileName(fileName);
            newFileName = newFileName.Substring(0, (newFileName.Length) - 4);
            newFileName = newFileName + ".xml";

            try
            {
                // execute the Microsoft Project Application
                app = new MSProject.ApplicationClass();
                // Do not display Microsoft Project
                app.Visible = false;
                // open the project file.
                if (app.FileOpen(fileName, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, MSProject.PjPoolOpen.pjPoolReadOnly, Type.Missing, Type.Missing, Type.Missing, Type.Missing))
                {
                    // go through all the open projects--there should only be one
                    foreach (MSProject.Project proj in app.Projects)
                    {
                        projects.Add(new ConsXMLExtract.ProjectContent(proj));

                        // and the most senior task of the project *************************
                        tasks.Add(new ConsXMLExtract.Task(proj.ProjectSummaryTask));

                        // go through all the tasks in the project
                        foreach (MSProject.Task task in proj.Tasks)
                        {
                            // we are only interested in tasks that do not have
                            // any child tasks--these are the tasks that we
                            // want to track.
                            //if (task.OutlineChildren.Count == 0)
                            {
                                // copy the Microsoft Project Task to our task
                                // and add it to our task list.
                                tasks.Add(new ConsXMLExtract.Task(task));
                            }
                        }
                    }
                }
                else
                {
                    retVal = "The MS Project file " + fileName + " could not be opened.";
                }
            }
            catch (Exception ex)
            {
                retVal = "Could not process the MS Project file " + fileName + "." + System.Environment.NewLine + ex.Message + System.Environment.NewLine + ex.StackTrace;
            }

            // close the application if is was opened.
            if (app != null)
            {
                app.Quit(MSProject.PjSaveType.pjDoNotSave);
            }
            return retVal;
        }

        /// <summary>
        /// Evalutes the tasks based upon day that you want to evalute the task
        /// from, and the number of days to look into the future.
        /// </summary>
        /// <param name="dayToEvaluateFrom">The day that you want to start evaluating the tasks from.</param>
        /// <param name="pastDays">The number of days to look into the past for the data to be considered current.</param>
        /// <param name="futureDays">The number of days that you want to look into the future.</param>
        //public void EvaluateTasks(System.DateTime dayToEvaluateFrom, int pastDays, int futureDays)
        public void EvaluateTasks()
        {
            //System.DateTime currentDate = dayToEvaluateFrom.AddDays(-pastDays);
            //System.DateTime futureDate = dayToEvaluateFrom.AddDays(futureDays);

            // clear out all of our data storage.
            late.Clear();
            current.Clear();
            future.Clear();
            currentProjects.Clear();

            foreach (ProjectContent project in projects)
            {
                currentProjects.Add(project);
            }


            // go through an evaluate each task
            foreach (Task task in tasks)
            {

                // make all task appear in current segment
                current.Add(task);


                // evaluate the late tasks
                //if (task.FinishDate <= dayToEvaluateFrom && task.PercentComplete != 100)
                //{
                //    late.Add(task);
                //}

                // evaluate the current tasks. we have for scenarios
                // 1. the task started after currentDate but before datToEvaluateFrom
                // 2. the task finished after currentDate but before datToEvaluateFrom
                // 3. the task started before currentDate and ends after the dayToEvaluateFrom
                //if ((task.StartDate >= currentDate && task.StartDate <= dayToEvaluateFrom) || // 1
                //    (task.FinishDate >= currentDate && task.FinishDate <= dayToEvaluateFrom) || // 2
                //    (task.StartDate < currentDate && task.FinishDate > dayToEvaluateFrom))
                //{ // 3
                //    current.Add(task);
                //}

                // evaluate the future tasks
                //if (task.StartDate > dayToEvaluateFrom && task.StartDate <= futureDate)
                //{
                //    future.Add(task);
                //}
            }
        }


        /// <summary>
        /// Saves the data to the XML file.
        /// </summary>
        /// <param name="fileName">The full path name of the file to save to.</param>
        /// <returns>On success an empty string; on error the error message.</returns>
        public string Save(string fileName)
        {
            System.IO.StreamWriter writer = null;
            string retVal = "";



            try
            {
                writer = new System.IO.StreamWriter(fileName, false, System.Text.Encoding.ASCII);
                writer.WriteLine(XmlPart1());


                foreach (ProjectContent project in currentProjects)
                {
                    writer.WriteLine(project.ToXml());
                }

                writer.WriteLine("<Tasks>");
                foreach (Task task in current)
                {
                    writer.WriteLine(task.ToXml());
                }
                writer.WriteLine("</Tasks>");
                //writer.WriteLine(XmlPart2());
                //foreach (Task task in late)
                //{
                //    writer.WriteLine(task.ToXml());
                //}
                writer.WriteLine(XmlPart3());       // add resource node segment
                //foreach (Task task in future)
                //{
                //    writer.WriteLine(task.ToXml());
                //}
                writer.WriteLine(XmlPart4());  // end of project and root nodes
            }
            catch (Exception ex)
            {
                retVal = "Could not save the file to " + fileName + "." + System.Environment.NewLine + ex.Message + System.Environment.NewLine + ex.StackTrace;
            }
            finally
            {
                if (writer != null)
                {
                    writer.Close();
                }
            }
            return retVal;
        }

        #region XML
        /// <summary>
        /// Returns the first part of the Excel XML file.  All the XML up to the
        /// items in the Current wrok sheet.
        /// </summary>
        /// <returns>The first part of the Excel XML file.</returns>
        private string XmlPart1()
        {
            string s = "";
            s += "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + System.Environment.NewLine;
            s += "<root>" + System.Environment.NewLine;
            s += "<Project>";


            //s += "<?mso-application progid=\"Excel.Sheet\"?>" + System.Environment.NewLine;
            //s += "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"" + System.Environment.NewLine;
            //s += "xmlns:o=\"urn:schemas-microsoft-com:office:office\"" + System.Environment.NewLine;
            //s += "xmlns:x=\"urn:schemas-microsoft-com:office:excel\"" + System.Environment.NewLine;
            //s += "xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"" + System.Environment.NewLine;
            //s += "xmlns:html=\"http://www.w3.org/TR/REC-html40\">" + System.Environment.NewLine;
            //s += "<DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\">" + System.Environment.NewLine;
            //s += "<Author>josh.carrillo@hpe.com</Author>" + System.Environment.NewLine;
            //s += "<LastAuthor>josh.carrillo@hpe.com</LastAuthor>" + System.Environment.NewLine;
            //s += "<Created>" + DateTime.UtcNow + "</Created>" + System.Environment.NewLine;
            //s += "<LastSaved>" + DateTime.UtcNow + "</LastSaved>" + System.Environment.NewLine;
            //s += "<Company></Company>" + System.Environment.NewLine;
            //s += "<Version>11.8107</Version>" + System.Environment.NewLine;
            //s += "</DocumentProperties>" + System.Environment.NewLine;
            //s += "<ExcelWorkbook xmlns=\"urn:schemas-microsoft-com:office:excel\">" + System.Environment.NewLine;
            //s += "<WindowHeight>8700</WindowHeight>" + System.Environment.NewLine;
            //s += "<WindowWidth>15195</WindowWidth>" + System.Environment.NewLine;
            //s += "<WindowTopX>480</WindowTopX>" + System.Environment.NewLine;
            //s += "<WindowTopY>135</WindowTopY>" + System.Environment.NewLine;
            //s += "<ProtectStructure>False</ProtectStructure>" + System.Environment.NewLine;
            //s += "<ProtectWindows>False</ProtectWindows>" + System.Environment.NewLine;
            //s += "</ExcelWorkbook>" + System.Environment.NewLine;
            //s += "<Styles>" + System.Environment.NewLine;
            //s += "<Style ss:ID=\"Default\" ss:Name=\"Normal\">" + System.Environment.NewLine;
            //s += "<Alignment ss:Vertical=\"Bottom\"/>" + System.Environment.NewLine;
            //s += "<Borders/>" + System.Environment.NewLine;
            //s += "<Font/>" + System.Environment.NewLine;
            //s += "<Interior/>" + System.Environment.NewLine;
            //s += "<NumberFormat/>" + System.Environment.NewLine;
            //s += "<Protection/>" + System.Environment.NewLine;
            //s += "</Style>" + System.Environment.NewLine;
            //s += "<Style ss:ID=\"s20\" ss:Name=\"Percent\">" + System.Environment.NewLine;
            //s += "<NumberFormat ss:Format=\"0%\"/>" + System.Environment.NewLine;
            //s += "</Style>" + System.Environment.NewLine;
            //s += "<Style ss:ID=\"s21\">" + System.Environment.NewLine;
            //s += "<NumberFormat ss:Format=\"Short Date\"/>" + System.Environment.NewLine;
            //s += "</Style>" + System.Environment.NewLine;
            //s += "<Style ss:ID=\"s27\">" + System.Environment.NewLine;
            //s += "<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/>" + System.Environment.NewLine;
            //s += "<Font x:Family=\"Swiss\" ss:Bold=\"1\"/>" + System.Environment.NewLine;
            //s += "<Interior ss:Color=\"#C0C0C0\" ss:Pattern=\"Solid\"/>" + System.Environment.NewLine;
            //s += "</Style>" + System.Environment.NewLine;
            //s += "<Style ss:ID=\"s28\" ss:Parent=\"s20\">" + System.Environment.NewLine;
            //s += "<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/>" + System.Environment.NewLine;
            //s += "<Font x:Family=\"Swiss\" ss:Bold=\"1\"/>" + System.Environment.NewLine;
            //s += "<Interior ss:Color=\"#C0C0C0\" ss:Pattern=\"Solid\"/>" + System.Environment.NewLine;
            //s += "</Style>" + System.Environment.NewLine;
            //s += "</Styles>" + System.Environment.NewLine;
            //s += "<Worksheet ss:Name=\"Current\">" + System.Environment.NewLine;
            //s += "<Names>" + System.Environment.NewLine;
            //s += "<NamedRange ss:Name=\"Print_Titles\" ss:RefersTo=\"=Current!R1\"/>" + System.Environment.NewLine;
            //s += "</Names>" + System.Environment.NewLine;
            //s += "<Table ss:ExpandedColumnCount=\"5\" x:FullColumns=\"1\"" + System.Environment.NewLine;
            //s += "x:FullRows=\"1\">" + System.Environment.NewLine;
            //s += "<Column ss:AutoFitWidth=\"0\" ss:Width=\"96\"/>" + System.Environment.NewLine;
            //s += "<Column ss:AutoFitWidth=\"0\" ss:Width=\"340.5\"/>" + System.Environment.NewLine;
            //s += "<Column ss:AutoFitWidth=\"0\" ss:Width=\"66.75\"/>" + System.Environment.NewLine;
            //s += "<Column ss:AutoFitWidth=\"0\" ss:Width=\"66\"/>" + System.Environment.NewLine;
            //s += "<Column ss:StyleID=\"s20\" ss:AutoFitWidth=\"0\" ss:Width=\"59.25\"/>" + System.Environment.NewLine;
            //s += "<Row ss:StyleID=\"s27\">" + System.Environment.NewLine;
            //s += "<Cell><Data ss:Type=\"String\">WBS</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>" + System.Environment.NewLine;
            //s += "<Cell><Data ss:Type=\"String\">Name</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>" + System.Environment.NewLine;
            //s += "<Cell><Data ss:Type=\"String\">Start Date</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>" + System.Environment.NewLine;
            //s += "<Cell><Data ss:Type=\"String\">Finish Date</Data><NamedCell" + System.Environment.NewLine;
            //s += "ss:Name=\"Print_Titles\"/></Cell>" + System.Environment.NewLine;
            //s += "<Cell ss:StyleID=\"s28\"><Data ss:Type=\"String\">% Complete</Data><NamedCell" + System.Environment.NewLine;
            //s += "ss:Name=\"Print_Titles\"/></Cell>" + System.Environment.NewLine;
            //s += "</Row>";



            return s;

        }

        /// <summary>
        /// Returns the second part of the Excel XML file.  All the XML after the
        /// current items up to the late items.
        /// </summary>
        /// <returns>The second part of the Excel XML file.</returns>
        private string XmlPart2()
        {
            string s = "";
            s += "</Table>" + System.Environment.NewLine;
            s += "<WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">" + System.Environment.NewLine;
            s += "<PageSetup>" + System.Environment.NewLine;
            s += "<Layout x:Orientation=\"Landscape\"/>" + System.Environment.NewLine;
            s += "<Header x:Data=\"&amp;A\"/>" + System.Environment.NewLine;
            s += "<Footer x:Data=\"Page &amp;P of &amp;N\"/>" + System.Environment.NewLine;
            s += "</PageSetup>" + System.Environment.NewLine;
            s += "<Print>" + System.Environment.NewLine;
            s += "<ValidPrinterInfo/>" + System.Environment.NewLine;
            s += "<HorizontalResolution>1200</HorizontalResolution>" + System.Environment.NewLine;
            s += "<VerticalResolution>1200</VerticalResolution>" + System.Environment.NewLine;
            s += "</Print>" + System.Environment.NewLine;
            s += "<Selected/>" + System.Environment.NewLine;
            s += "<Panes>" + System.Environment.NewLine;
            s += "<Pane>" + System.Environment.NewLine;
            s += "<Number>3</Number>" + System.Environment.NewLine;
            s += "<ActiveRow>1</ActiveRow>" + System.Environment.NewLine;
            s += "</Pane>" + System.Environment.NewLine;
            s += "</Panes>" + System.Environment.NewLine;
            s += "<ProtectObjects>False</ProtectObjects>" + System.Environment.NewLine;
            s += "<ProtectScenarios>False</ProtectScenarios>" + System.Environment.NewLine;
            s += "</WorksheetOptions>" + System.Environment.NewLine;
            s += "</Worksheet>" + System.Environment.NewLine;
            s += "<Worksheet ss:Name=\"Late\">" + System.Environment.NewLine;
            s += "<Names>" + System.Environment.NewLine;
            s += "<NamedRange ss:Name=\"Print_Titles\" ss:RefersTo=\"=Late!R1\"/>" + System.Environment.NewLine;
            s += "</Names>" + System.Environment.NewLine;
            s += "<Table ss:ExpandedColumnCount=\"5\" x:FullColumns=\"1\"" + System.Environment.NewLine;
            s += "x:FullRows=\"1\">" + System.Environment.NewLine;
            s += "<Column ss:AutoFitWidth=\"0\" ss:Width=\"96\"/>" + System.Environment.NewLine;
            s += "<Column ss:AutoFitWidth=\"0\" ss:Width=\"340.5\"/>" + System.Environment.NewLine;
            s += "<Column ss:AutoFitWidth=\"0\" ss:Width=\"66.75\"/>" + System.Environment.NewLine;
            s += "<Column ss:AutoFitWidth=\"0\" ss:Width=\"66\"/>" + System.Environment.NewLine;
            s += "<Column ss:StyleID=\"s20\" ss:AutoFitWidth=\"0\" ss:Width=\"59.25\"/>" + System.Environment.NewLine;
            s += "<Row ss:StyleID=\"s27\">" + System.Environment.NewLine;
            s += "<Cell><Data ss:Type=\"String\">WBS</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>" + System.Environment.NewLine;
            s += "<Cell><Data ss:Type=\"String\">Name</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>" + System.Environment.NewLine;
            s += "<Cell><Data ss:Type=\"String\">Start Date</Data><NamedCell ss:Name=\"Print_Titles\"/></Cell>" + System.Environment.NewLine;
            s += "<Cell><Data ss:Type=\"String\">Finish Date</Data><NamedCell" + System.Environment.NewLine;
            s += "ss:Name=\"Print_Titles\"/></Cell>" + System.Environment.NewLine;
            s += "<Cell ss:StyleID=\"s28\"><Data ss:Type=\"String\">% Complete</Data><NamedCell" + System.Environment.NewLine;
            s += "ss:Name=\"Print_Titles\"/></Cell>" + System.Environment.NewLine;
            s += "</Row>";
            return s;
        }

        /// <summary>
        /// Returns the third part of the Excel XML file.  All the XML after the
        /// late items up to the future items.
        /// </summary>
        /// <returns>The third part of the Excel XML file.</returns>
        private string XmlPart3()
        {
            string s = "";

            s += "<Resources>" + System.Environment.NewLine;
            s += "<Resource>" + System.Environment.NewLine;
            s += "<UID>0</UID>" + System.Environment.NewLine;
            s += "<ID>0</ID>" + System.Environment.NewLine;
            s += "<Type>1</Type>" + System.Environment.NewLine;
            s += "<IsNull>0</IsNull>" + System.Environment.NewLine;
            s += "<WorkGroup>0</WorkGroup>" + System.Environment.NewLine;
            s += "<MaxUnits>1.00</MaxUnits>" + System.Environment.NewLine;
            s += "<PeakUnits>0.00</PeakUnits>" + System.Environment.NewLine;
            s += "<OverAllocated>0</OverAllocated>" + System.Environment.NewLine;
            s += "<CanLevel>1</CanLevel>" + System.Environment.NewLine;
            s += "<AccrueAt>3</AccrueAt>" + System.Environment.NewLine;
            s += "<Work>PT0H0M0S</Work>" + System.Environment.NewLine;
            s += "<RegularWork>PT0H0M0S</RegularWork>" + System.Environment.NewLine;
            s += "<OvertimeWork>PT0H0M0S</OvertimeWork>" + System.Environment.NewLine;
            s += "<ActualWork>PT0H0M0S</ActualWork>" + System.Environment.NewLine;
            s += "<RemainingWork>PT0H0M0S</RemainingWork>" + System.Environment.NewLine;
            s += "<ActualOvertimeWork>PT0H0M0S</ActualOvertimeWork>" + System.Environment.NewLine;
            s += "<RemainingOvertimeWork>PT0H0M0S</RemainingOvertimeWork>" + System.Environment.NewLine;
            s += "<PercentWorkComplete>0</PercentWorkComplete>" + System.Environment.NewLine;
            s += "<StandardRate>0</StandardRate>" + System.Environment.NewLine;
            s += "<StandardRateFormat>2</StandardRateFormat>" + System.Environment.NewLine;
            s += "<Cost>0</Cost>" + System.Environment.NewLine;
            s += "<OvertimeRate>0</OvertimeRate>" + System.Environment.NewLine;
            s += "<OvertimeRateFormat>2</OvertimeRateFormat>" + System.Environment.NewLine;
            s += "<OvertimeCost>0</OvertimeCost>" + System.Environment.NewLine;
            s += "<CostPerUse>0</CostPerUse>" + System.Environment.NewLine;
            s += "<ActualCost>0</ActualCost>" + System.Environment.NewLine;
            s += "<ActualOvertimeCost>0</ActualOvertimeCost>" + System.Environment.NewLine;
            s += "<RemainingCost>0</RemainingCost>" + System.Environment.NewLine;
            s += "<RemainingOvertimeCost>0</RemainingOvertimeCost>" + System.Environment.NewLine;
            s += "<WorkVariance>0.00</WorkVariance>" + System.Environment.NewLine;
            s += "<CostVariance>0</CostVariance>" + System.Environment.NewLine;
            s += "<SV>0.00</SV>" + System.Environment.NewLine;
            s += "<CV>0.00</CV>" + System.Environment.NewLine;
            s += "<ACWP>0.00</ACWP>" + System.Environment.NewLine;
            s += "<CalendarUID>2</CalendarUID>" + System.Environment.NewLine;
            s += "<BCWS>0.00</BCWS>" + System.Environment.NewLine;
            s += "<BCWP>0.00</BCWP>" + System.Environment.NewLine;
            s += "<IsGeneric>0</IsGeneric>" + System.Environment.NewLine;
            s += "<IsInactive>0</IsInactive>" + System.Environment.NewLine;
            s += "<IsEnterprise>0</IsEnterprise>" + System.Environment.NewLine;
            s += "<BookingType>0</BookingType>" + System.Environment.NewLine;
            s += "<CreationDate>2017-09-15T12:15:00</CreationDate>" + System.Environment.NewLine;
            s += "<IsCostResource>0</IsCostResource>" + System.Environment.NewLine;
            s += "<IsBudget>0</IsBudget>" + System.Environment.NewLine;
            s += "</Resource>" + System.Environment.NewLine;
            s += "</Resources>" + System.Environment.NewLine;
            s += "<Assignments>" + System.Environment.NewLine;
            s += "<Assignment>" + System.Environment.NewLine;
            s += "<UID>9</UID>" + System.Environment.NewLine;
            s += "<TaskUID>9</TaskUID>" + System.Environment.NewLine;
            s += "<ResourceUID>-65535</ResourceUID>" + System.Environment.NewLine;
            s += "<PercentWorkComplete>50</PercentWorkComplete>" + System.Environment.NewLine;
            s += "<ActualCost>0</ActualCost>" + System.Environment.NewLine;
            s += "<ActualOvertimeCost>0</ActualOvertimeCost>" + System.Environment.NewLine;
            s += "<ActualOvertimeWork>PT0H0M0S</ActualOvertimeWork>" + System.Environment.NewLine;
            s += "<ActualStart>2018-01-26T08:00:00</ActualStart>" + System.Environment.NewLine;
            s += "<ActualWork>PT4H0M0S</ActualWork>" + System.Environment.NewLine;
            s += "<ACWP>0.00</ACWP>" + System.Environment.NewLine;
            s += "<Confirmed>0</Confirmed>" + System.Environment.NewLine;
            s += "<Cost>0</Cost>" + System.Environment.NewLine;
            s += "<CostRateTable>0</CostRateTable>" + System.Environment.NewLine;
            s += "<RateScale>0</RateScale>" + System.Environment.NewLine;
            s += "<CostVariance>0</CostVariance>" + System.Environment.NewLine;
            s += "<CV>-0.00</CV>" + System.Environment.NewLine;
            s += "<Delay>0</Delay>" + System.Environment.NewLine;
            s += "<Finish>2018-01-26T17:00:00</Finish>" + System.Environment.NewLine;
            s += "<FinishVariance>0</FinishVariance>" + System.Environment.NewLine;
            s += "<WorkVariance>0.00</WorkVariance>" + System.Environment.NewLine;
            s += "<HasFixedRateUnits>1</HasFixedRateUnits>" + System.Environment.NewLine;
            s += "<FixedMaterial>0</FixedMaterial>" + System.Environment.NewLine;
            s += "<LevelingDelay>0</LevelingDelay>" + System.Environment.NewLine;
            s += "<LevelingDelayFormat>7</LevelingDelayFormat>" + System.Environment.NewLine;
            s += "<LinkedFields>0</LinkedFields>" + System.Environment.NewLine;
            s += "<Milestone>0</Milestone>" + System.Environment.NewLine;
            s += "<Overallocated>0</Overallocated>" + System.Environment.NewLine;
            s += "<OvertimeCost>0</OvertimeCost>" + System.Environment.NewLine;
            s += "<OvertimeWork>PT0H0M0S</OvertimeWork>" + System.Environment.NewLine;
            s += "<RegularWork>PT8H0M0S</RegularWork>" + System.Environment.NewLine;
            s += "<RemainingCost>0</RemainingCost>" + System.Environment.NewLine;
            s += "<RemainingOvertimeCost>0</RemainingOvertimeCost>" + System.Environment.NewLine;
            s += "<RemainingOvertimeWork>PT0H0M0S</RemainingOvertimeWork>" + System.Environment.NewLine;
            s += "<RemainingWork>PT4H0M0S</RemainingWork>" + System.Environment.NewLine;
            s += "<ResponsePending>0</ResponsePending>" + System.Environment.NewLine;
            s += "<Start>2018-01-26T08:00:00</Start>" + System.Environment.NewLine;
            s += "<Stop>2018-01-26T12:00:00</Stop>" + System.Environment.NewLine;
            s += "<Resume>2018-01-26T13:00:00</Resume>" + System.Environment.NewLine;
            s += "<StartVariance>0</StartVariance>" + System.Environment.NewLine;
            s += "<Units>1</Units>" + System.Environment.NewLine;
            s += "<UpdateNeeded>0</UpdateNeeded>" + System.Environment.NewLine;
            s += "<VAC>0.00</VAC>" + System.Environment.NewLine;
            s += "<Work>PT8H0M0S</Work>" + System.Environment.NewLine;
            s += "<WorkContour>0</WorkContour>" + System.Environment.NewLine;
            s += "<BCWS>0.00</BCWS>" + System.Environment.NewLine;
            s += "<BCWP>-0.00</BCWP>" + System.Environment.NewLine;
            s += "<BookingType>0</BookingType>" + System.Environment.NewLine;
            s += "<CreationDate>2017-09-15T12:15:00</CreationDate>" + System.Environment.NewLine;
            s += "<BudgetCost>0</BudgetCost>" + System.Environment.NewLine;
            s += "<BudgetWork>PT0H0M0S</BudgetWork>" + System.Environment.NewLine;
            s += "<TimephasedData>" + System.Environment.NewLine;
            s += "<Type>2</Type>" + System.Environment.NewLine;
            s += "<UID>9</UID>" + System.Environment.NewLine;
            s += "<Start>2018-01-26T08:00:00</Start>" + System.Environment.NewLine;
            s += "<Finish>2018-01-26T12:00:00</Finish>" + System.Environment.NewLine;
            s += "<Unit>1</Unit>" + System.Environment.NewLine;
            s += "<Value>PT4H0M0S</Value>" + System.Environment.NewLine;
            s += "</TimephasedData>" + System.Environment.NewLine;
            s += "</Assignment>" + System.Environment.NewLine;
            s += "</Assignments>";



            return s;
        }

        /// <summary>
        /// Returns the fourth part of the Excel XML file.  All the XML after the
        /// future items to the end of the file.
        /// </summary>
        /// <returns>The second part of the Excel XML file.</returns>
        public string XmlPart4()
        {
            string s = "";

            s += "</Project>" + System.Environment.NewLine;
            s += "</root>" + System.Environment.NewLine;

            //s += "</Table>" + System.Environment.NewLine;
            //s += "<WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">" + System.Environment.NewLine;
            //s += "<PageSetup>" + System.Environment.NewLine;
            //s += "<Layout x:Orientation=\"Landscape\"/>" + System.Environment.NewLine;
            //s += "<Header x:Data=\"&amp;A\"/>" + System.Environment.NewLine;
            //s += "<Footer x:Data=\"Page &amp;P of &amp;N\"/>" + System.Environment.NewLine;
            //s += "</PageSetup>" + System.Environment.NewLine;
            //s += "<Print>" + System.Environment.NewLine;
            //s += "<ValidPrinterInfo/>" + System.Environment.NewLine;
            //s += "<HorizontalResolution>1200</HorizontalResolution>" + System.Environment.NewLine;
            //s += "<VerticalResolution>1200</VerticalResolution>" + System.Environment.NewLine;
            //s += "</Print>" + System.Environment.NewLine;
            //s += "<Panes>" + System.Environment.NewLine;
            //s += "<Pane>" + System.Environment.NewLine;
            //s += "<Number>3</Number>" + System.Environment.NewLine;
            //s += "<ActiveRow>1</ActiveRow>" + System.Environment.NewLine;
            //s += "<ActiveCol>1</ActiveCol>" + System.Environment.NewLine;
            //s += "</Pane>" + System.Environment.NewLine;
            //s += "</Panes>" + System.Environment.NewLine;
            //s += "<ProtectObjects>False</ProtectObjects>" + System.Environment.NewLine;
            //s += "<ProtectScenarios>False</ProtectScenarios>" + System.Environment.NewLine;
            //s += "</WorksheetOptions>" + System.Environment.NewLine;
            //s += "</Worksheet>" + System.Environment.NewLine;
            //s += "</Workbook>" + System.Environment.NewLine;


            return s;
        }
        #endregion XML



    }
}
