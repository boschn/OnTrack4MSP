using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using LiteDB;
using Microsoft.SqlServer.Server;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;

namespace OnTrackMSP
{

    /// <summary>
    /// an Excel Gantt representation
    /// </summary>
    class PublishExcelGantt
    {
        /// <summary>
        /// internal class to define a column type
        /// </summary>
        public enum ColumnType
        {
            Taskname = 1,
            TaskUID,
            TaskID,
            TaskProgress,
            TaskSBSCode,
            TaskStart,
            TaskFinish,
            TaskOutlineLevel,
            TaskBaselinecode,
            TaskDelegated,
            TaskPlanType,
            TaskProductRelease,
            TaskResources,
            TaskVariantName,
            TaskPredecessor,
            TaskResponsible,
            TaskOutlineChildren
        }


        /// <summary>
        /// internal class to define a column of the table
        /// </summary>
        public class TableColumnDef
        {
            public ColumnType Type { get; set; }
            public string Title { get; set; }
            public uint ColumnNumber { get; set; }
            public string Format { get; set; }
            public ExcelStyle HeaderStyle { get; set; }

            public ExcelStyle ValueStyle { get; set; }


            /// <summary>
            /// ctor
            /// </summary>
            /// <param name="type"></param>
            /// <param name="Title"></param>
            /// <param name="ColumnNumber"></param>
            /// <param name="Format"></param>
            public TableColumnDef(ColumnType type, string Title, uint ColumnNumber, string Format = null, ExcelStyle HeaderStyle = null)

            {
                this.Type = type;
                this.Title = Title;
                this.ColumnNumber = ColumnNumber;
                this.Format = Format;
                this.HeaderStyle = HeaderStyle;

            }
        }


        /// <summary>
        /// Definition of the Table for Roadmaps
        /// </summary>
        public  static PublishExcelGantt.TableColumnDef[] RoadmapTableDef = {
            new TableColumnDef(PublishExcelGantt.ColumnType.TaskUID, "MSPUID", 1),
            new TableColumnDef(PublishExcelGantt.ColumnType.Taskname, "Name", 2),
            new TableColumnDef(PublishExcelGantt.ColumnType.TaskProgress, "Progress", 3),
            new TableColumnDef(PublishExcelGantt.ColumnType.TaskStart, "Start", 4, "dd.MM.yyyy"),
            new TableColumnDef(PublishExcelGantt.ColumnType.TaskFinish, "Finish", 5, "dd.MM.yyyy"),
            new TableColumnDef(PublishExcelGantt.ColumnType.TaskResources, "Resources", 6),
            new TableColumnDef(PublishExcelGantt.ColumnType.TaskSBSCode, "SBSCode", 7)

        };
        /// <summary>
        /// Definition of the Table for BaselineOverview
        /// </summary>
        public static PublishExcelGantt.TableColumnDef[] BaselineOverViewTableDef = {
            new TableColumnDef(PublishExcelGantt.ColumnType.TaskUID, "MSPUID", 1),
            new TableColumnDef(PublishExcelGantt.ColumnType.TaskBaselinecode, "Baseline", 2),
            new TableColumnDef(PublishExcelGantt.ColumnType.Taskname, "Name", 3),
            new TableColumnDef(PublishExcelGantt.ColumnType.TaskFinish, "Date", 4, "dd.MM.yyyy"),
            new TableColumnDef(PublishExcelGantt.ColumnType.TaskResponsible, "Responsible", 5),
            new TableColumnDef(PublishExcelGantt.ColumnType.TaskProductRelease, "ProductRelease", 6),
            new TableColumnDef(PublishExcelGantt.ColumnType.TaskVariantName, "Variant", 7),
        };

        // global definitions of the Gantt Class
        private string ganttName { get; }
        private ExcelRange ganttRange;
        private IList<dbTask> tasklist;
        public uint MaxRow { get; internal set; }
        public DateTime? startGanttDate { get; set; }
        public DateTime? endGanttDate { get; set; }
        private Calendar validCalendar { get; set; }
      
        public IList<TableColumnDef> TableDefinition { get; internal set; }
        private ExcelStyle tableHeaderStyle { get; set; }
        private ExcelNamedStyleXml ganttTableHeaderStyle { get; set; }
        private ExcelNamedStyleXml ganttChartHeaderStyle { get; set; }

        /// <summary>
        /// ctor a gantt chart object for the range and the tasklist
        /// </summary>
        /// <param name="ganttRange"></param>
        /// <param name="tasklist"></param>
        public PublishExcelGantt(string name, ExcelRange ganttRange, IList<dbTask> tasklist)
        {
            // Name it
            if (String.IsNullOrEmpty(name)) ganttName = ganttRange.Worksheet.Name + "_gantt";
            else ganttName = name;
            // range
            this.ganttRange = ganttRange;
            // tasklist
            this.tasklist = tasklist;

            // default generate dates
            if (this.startGanttDate== null)
                this.startGanttDate =  new DateTime(DateTime.Now.Year, 1, 1);
            if (this.endGanttDate == null)
                this.endGanttDate = new DateTime(DateTime.Now.Year+2, 12, 31);
            
            // default
            _ = (DateTimeFormatInfo.CurrentInfo != null) ? this.validCalendar = DateTimeFormatInfo.CurrentInfo.Calendar : this.validCalendar = new GregorianCalendar();

           // default header style
            ganttTableHeaderStyle =
                ganttRange.Worksheet.Workbook.Styles.NamedStyles.FirstOrDefault(x =>
                    x.Name == ganttName + @"_tableheaderstyle");
            if (ganttTableHeaderStyle == null)
            {
                ganttTableHeaderStyle = ganttRange.Worksheet.Workbook.Styles.CreateNamedStyle(ganttName + @"_tableheaderstyle");
                ganttTableHeaderStyle.Style.Font.Bold = true;
                ganttTableHeaderStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
                ganttTableHeaderStyle.Style.Fill.BackgroundColor.SetColor(Color.DarkGray);

                ganttTableHeaderStyle.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                ganttTableHeaderStyle.Style.Border.Bottom.Color.SetColor(Color.Black);
                ganttTableHeaderStyle.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ganttTableHeaderStyle.Style.Border.Top.Color.SetColor(Color.Black);
            }
            // default header style
            ganttChartHeaderStyle =
                ganttRange.Worksheet.Workbook.Styles.NamedStyles.FirstOrDefault(x =>
                    x.Name == ganttName + @"_chartheaderstyle");
            if (ganttChartHeaderStyle == null)
            {
                ganttChartHeaderStyle = ganttRange.Worksheet.Workbook.Styles.CreateNamedStyle(ganttName + @"_chartheaderstyle");
                ganttChartHeaderStyle.Style.Font.Bold = false;
                ganttChartHeaderStyle.Style.TextRotation = 180;
                ganttChartHeaderStyle.Style.Font.Size = 06;
                ganttChartHeaderStyle.Style.Font.Color.SetColor(Color.Blue);
                ganttChartHeaderStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
                ganttChartHeaderStyle.Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);
                ganttChartHeaderStyle.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                ganttChartHeaderStyle.Style.Border.Bottom.Color.SetColor(Color.Black);
                ganttChartHeaderStyle.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ganttChartHeaderStyle.Style.Border.Top.Color.SetColor(Color.Black);
                
            }
        }

        /// <summary>
        /// returns the number of calendar weeks between 2 dates
        /// </summary>
        /// <param name="fromDate"></param>
        /// <param name="toDate"></param>
        /// <returns></returns>
        private int GetNumberofWeeks(DateTime fromDate, DateTime toDate)
        {
            return (validCalendar.GetWeekOfYear(toDate, CalendarWeekRule.FirstDay, DayOfWeek.Monday) -
                    validCalendar.GetWeekOfYear(fromDate, CalendarWeekRule.FirstDay, DayOfWeek.Monday));
        }
        /// <summary>
        /// calculate the column index for a date
        /// return negativ if earlier
        /// </summary>
        /// <param name="task"></param>
        private int GetGanttColumn(DateTime date)
        {
            int aWeek = 0;
            int aYear = startGanttDate.Value.Year;
            // return negativ if earlier
            if (aYear > date.Year) return -1;

            // sum up the weaks
            while (aYear < date.Year)
            {
                aWeek += GetNumberofWeeks(new DateTime(aYear, 1, 1), new DateTime(aYear, 12, 31));
                aYear++;
            }
            // add the current number of weeks
            aWeek += validCalendar.GetWeekOfYear(date, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
            // return total
            return aWeek;
        }
        /// <summary>
        /// Generate the gantt in the range for the given tasks
        /// </summary>
        /// <returns></returns>
        public bool Generate(IList<TableColumnDef> TableDefinition)
        {
            
            var aTableEndCol = TableDefinition.Max(x => x.ColumnNumber);

            this.TableDefinition = TableDefinition;

            GenerateTableHeader(2,1);
            GenerateGanttHeader(1, (int)aTableEndCol + 1);
            this.MaxRow = 3;

            // for all tasks in tasklist
            foreach (var aTask in tasklist)
            {
                // only active & not in the past
                if (aTask.Finish.HasValue)
                    if (aTask.Finish > this.startGanttDate) 
                    {
                        // generate Table
                        GenerateTableRow((int)this.MaxRow, 1, aTask);

                        // generate Gantt Chart
                        GenerateGanttRow((int)this.MaxRow, (int)aTableEndCol + 1, aTask);

                        MaxRow++;
                    }
               
            }

            var aRange = ganttRange.Worksheet.Cells[2, 1, (int) this.MaxRow - 1,
                (int) aTableEndCol + GetGanttColumn(endGanttDate.Value)];

            // autofit for table part
            ganttRange.Worksheet.Cells[2, 1, (int)this.MaxRow - 1,
                (int)aTableEndCol].AutoFitColumns();

            // freeze pane
            ganttRange.Worksheet.View.FreezePanes(3,(int)aTableEndCol+1);

            // save the whole as data table
            var aTableName = Regex.Replace("table_" + ganttName, @"\s+", "_");
            if (ganttRange.Worksheet.Tables[aTableName] != null)
                ganttRange.Worksheet.Tables.Delete(aTableName, false);
            
            
            ganttRange.Worksheet.Tables.Add(aRange, aTableName);
            ganttRange.Worksheet.Tables[aTableName].ShowFilter = false;
            return true;
        }

       

        /// <summary>
        /// create the header of the gantt chart
        /// </summary>
        private void GenerateGanttHeader(int row, int column)
        {
            int aStartCol = 0;
            int anEndCol = 0;
            DateTime aDate = startGanttDate.Value;
            
            
            // set Start
            aStartCol = GetGanttColumn(startGanttDate.Value);
            
            // underflow -> starts in start col
            if (aStartCol > GetGanttColumn(endGanttDate.Value)) aStartCol = GetGanttColumn(endGanttDate.Value);
 
            // set End
            anEndCol = GetGanttColumn(endGanttDate.Value);

            // overflow -> stops in last date
            if (anEndCol < GetGanttColumn(startGanttDate.Value)) anEndCol = GetGanttColumn(startGanttDate.Value);
            if (anEndCol > GetGanttColumn(endGanttDate.Value)) anEndCol = GetGanttColumn(endGanttDate.Value);

            // draw the year bar
            for (var aYear = aDate.Year; aYear <= endGanttDate.Value.Year; aYear++)
            {
                var aYearStartCol = GetGanttColumn(new DateTime(aYear, 01, 01));
                // if the column is already merged (as in another year) then increase
                if (ganttRange.Worksheet.Cells[row, column + aYearStartCol - 1].Merge == true) 
                    aYearStartCol++;

                var aYearEndCol = GetGanttColumn(new DateTime(aYear, 12, 31));
                ganttRange.Worksheet.Cells[row, column + aYearStartCol - 1].Value = aYear;
                // merge the year
                if (ganttRange.Worksheet.Cells[row, column + aYearStartCol - 1].Merge != true)
                    ganttRange.Worksheet.Cells[row, column + aYearStartCol - 1, row, column + aYearEndCol - 1].Merge = true;
                else
                {
                    if (ganttRange.Worksheet.Cells[row, column + aYearStartCol].Merge != true)
                        ganttRange.Worksheet.Cells[row, column + aYearStartCol , row, column + aYearEndCol - 1].Merge = true;
                }
                // setz size
                ganttRange.Worksheet.Cells[row, column + aYearStartCol - 1, row, column + aYearEndCol - 1].Style.Font
                    .Size = 10;
                ganttRange.Worksheet.Cells[row, column + aYearStartCol - 1, row, column + aYearEndCol - 1].Style.Font
                    .Bold = true;

                // iterate background
                ganttRange.Worksheet.Cells[row, column + aYearStartCol - 1, row, column + aYearEndCol - 1].Style
                    .Fill.PatternType = ExcelFillStyle.Solid;
                if (aYear % 2 == 0)
                    ganttRange.Worksheet.Cells[row, column + aYearStartCol - 1, row, column + aYearEndCol - 1].Style
                        .Fill.BackgroundColor.SetColor(Color.LightCyan);
                else ganttRange.Worksheet.Cells[row, column + aYearStartCol - 1, row, column + aYearEndCol - 1].Style
                    .Fill.BackgroundColor.SetColor(Color.Gainsboro);
            }
            this.MaxRow = Convert.ToUInt32(row); // set MaxRow
            ganttRange.Worksheet.Row((int)this.MaxRow).Height = 20;
            

            // draw the gantt week bar
            for (int i = aStartCol; i <= anEndCol; i++)
            {

                ganttRange[row+1, column +i -1].Value = String.Format( "'{0:00}-{1:yy}", validCalendar.GetWeekOfYear(aDate, CalendarWeekRule.FirstDay, DayOfWeek.Monday), aDate);
                ganttRange[row+1, column + i - 1].StyleName = ganttChartHeaderStyle.Name;
                ganttRange.Worksheet.Column(column + i - 1).Width = 3;
                ganttRange[row, column + i - 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                ganttRange[row, column + i - 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                // increase week
                aDate = aDate.AddDays(7);


            }
            this.MaxRow = Convert.ToUInt32(row) +1; // set MaxRow
            ganttRange.Worksheet.Row((int)this.MaxRow).Height = 30;
        }
        /// <summary>
        /// create the header of the table
        /// </summary>
        private void GenerateTableHeader(int row, int column)
        {
            foreach (var aColumnDef in this.TableDefinition)
            {
                ganttRange[row, column  + (int)aColumnDef.ColumnNumber -1].Value = aColumnDef.Title;
                ganttRange[row, column + (int)aColumnDef.ColumnNumber -1].StyleName = ganttTableHeaderStyle.Name;
               
            }
        }

        /// <summary>
        /// write a table row for a task
        /// </summary>
        /// <param name="range"></param>
        /// <param name="task"></param>
        private  bool GenerateTableRow(int row, int column, dbTask task)
        {

            foreach (var aColumnDef in this.TableDefinition)
            {
                var ourColumn = column + (int) aColumnDef.ColumnNumber -1;
                if (!String.IsNullOrEmpty(aColumnDef.Format))
                    ganttRange[row, ourColumn].Style.Numberformat.Format =aColumnDef.Format;

                switch (aColumnDef.Type)
                {
                    case ColumnType.TaskUID:
                        ganttRange[row, ourColumn].Value = task.MSPUID;
                        break;
                    case ColumnType.TaskID:
                        ganttRange[row, ourColumn].Value = task.MSPUID;
                        break;
                    case ColumnType.Taskname:
                        ganttRange[row, ourColumn].Value = task.Name;
                        break;
                    case ColumnType.TaskOutlineLevel:
                        ganttRange[row, ourColumn].Value = task.OutlineLevel;
                        break;
                    case ColumnType.TaskProgress:
                        ganttRange[row, ourColumn].Value = task.Progress;
                        break;
                    case ColumnType.TaskSBSCode:
                        ganttRange[row, ourColumn].Value = task.SBSCode;
                        break;
                    case ColumnType.TaskStart:
                        ganttRange[row, ourColumn].Value = task.Start;
                        break;
                    case ColumnType.TaskFinish:
                        ganttRange[row, ourColumn].Value = task.Finish;
                        break;
                    case ColumnType.TaskBaselinecode:
                        ganttRange[row, ourColumn].Value = task.BaselineCode;
                        break;
                    case ColumnType.TaskDelegated:
                        ganttRange[row, ourColumn].Value = task.Delegated;
                        break;
                    case ColumnType.TaskResponsible:
                        ganttRange[row, ourColumn].Value = task.Responsible;
                        break;
                    case ColumnType.TaskPlanType:
                        ganttRange[row, ourColumn].Value = task.PlanType;
                        break;
                    case ColumnType.TaskProductRelease:
                        ganttRange[row, ourColumn].Value = task.ProductRelease;
                        break;
                    case ColumnType.TaskVariantName:
                        ganttRange[row, ourColumn].Value = task.VariantName;
                        break;
                    case ColumnType.TaskResources:
                    { 
                        ganttRange[row, ourColumn].Value = task.Resources;
                        string aValue = "";
                        foreach (var anId in task.Resources)
                        {
                                if (!String.IsNullOrEmpty(aValue)) aValue += ",";
                                aValue += " " + aValue;
                        }
                        ganttRange[row, ourColumn].Value = aValue;
                    }
                        break;
                    case ColumnType.TaskPredecessor:
                    {
                        string aValue = "";
                        foreach (var anId in task.Predecessors)
                        {
                            var aTask = DBase.GetTask(anId);
                            if (aTask != null)
                            {
                                if (!String.IsNullOrEmpty(aValue)) aValue += ",";
                                aValue += " " + aTask.MSPUID;
                            }
                        }

                        ganttRange[row, ourColumn].Value = aValue;
                    }
                        break;
                    case ColumnType.TaskOutlineChildren:
                    {
                        string aValue = "";
                        foreach (var anId in task.OutlineChildren)
                        {
                            var aTask = DBase.GetTask(anId);
                            if (aTask != null)
                            {
                                if (!String.IsNullOrEmpty(aValue)) aValue += ",";
                                aValue += " " + aTask.MSPUID;
                            }

                        }

                        ganttRange[row, ourColumn].Value = aValue;
                    }
                        break;
                    default:
                        Debug.WriteLine("Definition enumeration not implemented in GenerateTableRow");
                        break;
                }
            }
            
            return true;
        }
        /// <summary>
        /// write a gantt chart row for a task
        /// </summary>
        /// <param name="range"></param>
        /// <param name="task"></param>
        private  bool GenerateGanttRow(int row, int col, dbTask task)
        {
            int aStartCol=0;
            int anEndCol = 0;
            Color aGantColor = Color.CornflowerBlue;
            Color aTextColor = Color.Black;

            // determine the start column
            if (task.Start.HasValue)
                aStartCol = GetGanttColumn(task.Start.Value);
            // take the end - milestone logic
            else if (task.Finish.HasValue)
                aStartCol = GetGanttColumn(task.Finish.Value);
            else
            {
                // no value in end then overflow
                aStartCol = GetGanttColumn(endGanttDate.Value) +1;
            }
            // underflow -> starts in start col
            if (aStartCol < GetGanttColumn(startGanttDate.Value)) aStartCol = GetGanttColumn(startGanttDate.Value);
            if (aStartCol > GetGanttColumn(endGanttDate.Value)) aStartCol = GetGanttColumn(endGanttDate.Value);

            // determine the end column
            if (task.Finish.HasValue)
                anEndCol = GetGanttColumn(task.Finish.Value);
            // overflow
            else anEndCol = GetGanttColumn(endGanttDate.Value) + 1;

            // overflow -> stops in last date
            if (anEndCol < GetGanttColumn(startGanttDate.Value)) anEndCol = GetGanttColumn(startGanttDate.Value);
            if (anEndCol > GetGanttColumn(endGanttDate.Value)) anEndCol = GetGanttColumn(endGanttDate.Value);

            // Color
            aGantColor = Color.CornflowerBlue; // default
            if (task.IsLocoNeeded)
            {
                aGantColor = Color.DarkGreen;
                aTextColor = Color.White;
            }

            if (task.IsMilestone)
            {
                aGantColor = Color.Blue;
                aTextColor = Color.White;
            }

            if (task.IsGovernance)
            {
                aGantColor = Color.BlueViolet;
                aTextColor = Color.White;
            }

            if (task.IsPendTime)
            {
                aGantColor = Color.Khaki;
                aTextColor = Color.Black;
            }

            if (task.IsXtrnDriven)
            {
                aGantColor = Color.Yellow;
                aTextColor = Color.Black;
            }

            if (task.SBSCode == @"30.10")
            {
                aGantColor = Color.Yellow;
                aTextColor = Color.Black;
            }

            // draw the gantt bar
            for (int i = aStartCol; i <= anEndCol; i++)
            {
                ganttRange[row, col-1+i].Value = " ";
                ganttRange[row, col - 1+i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ganttRange[row, col - 1+i].Style.Fill.BackgroundColor.SetColor(aGantColor);
            }
            // mark milestone
            //
            if (task.IsMilestone || task.IsGovernance)
            {
                string aMilestoneName;

                // add the rollup
                if (!String.IsNullOrEmpty(task.RollupName)) aMilestoneName = "'" + task.RollupName;
                else aMilestoneName = "'MS";

                ganttRange[row, col + anEndCol - 1].Value = aMilestoneName;
                ganttRange[row, col + anEndCol - 1].Style.Font.Color.SetColor(aTextColor);
                ganttRange[row, col + anEndCol - 1].Style.Font.Size = 6;
                ganttRange[row, col + anEndCol - 1].Style.Font.Bold = true;
                ganttRange[row, col + anEndCol - 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ganttRange[row, col + anEndCol - 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                // add comment with details
                if (ganttRange[row, col + anEndCol - 1].Comment == null)
                {
                    ganttRange[row, col + anEndCol - 1]
                        .AddComment(
                            "Milestone " + aMilestoneName + " from " + task.MSPUID + " '" +
                            task.Name + string.Format("' finish on {0:yyyy-MM-dd}",
                                task.Finish.Value),
                            $"ontrack on {DateTime.Now.ToString("yyyy-MM-dd")}");
                    ganttRange[row, col + anEndCol - 1].Comment.AutoFit = true;
                }
                else
                {
                    ganttRange[row, col + anEndCol - 1].Comment.Text =
                        "Milestone " + aMilestoneName + "from " + task.MSPUID + " '" + task.Name +
                        string.Format("' finish on {0:yyyy-MM-dd}", task.Finish.Value);

                    ganttRange[row, col + anEndCol - 1].Comment.Author =
                        $"OnTrackTool on {DateTime.Now.ToString("yyyy-MM-dd")}";

                    ganttRange[row, col + anEndCol - 1].Comment.AutoFit = true;
                }
            }

            // draw rollups
            if (task.IsSummary && task.IsRollup)
                GenerateGanttRowRollup(row, col, task);
            return true;
        }

        /// <summary>
        /// generate a gantt row for a rollup in the same projectId
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="task"></param>
        /// <returns></returns>
        private bool GenerateGanttRowRollup(int row, int col, dbTask task)
        {
                // crawl through all sub tasks
                foreach (var aTaskID in task.OutlineChildren)
                {
                    // return task
                    
                    var aTask = DBase.GetTask(aTaskID);
                    if (aTask != null) 
                    {
                        // if task is rollup and no summary
                        if (aTask.IsRollup & !aTask.IsSummary)
                            if (aTask.Finish.HasValue)
                             
                            { 
                                var i = GetGanttColumn(aTask.Finish.Value);
                                // check if in the gantt
                                if (i >= GetGanttColumn(startGanttDate.Value) && i <= GetGanttColumn(endGanttDate.Value))
                                {
                                    // add the rollup
                                    ganttRange[row, col + i - 1].Value = aTask.RollupName;
                                    ganttRange[row, col + i - 1].Style.Font.Size = 6;
                                    ganttRange[row, col + i - 1].Style.Font.Bold =true;
                                    // add comment with details
                                    if (ganttRange[row, col + i - 1].Comment == null)
                                        {
                                            ganttRange[row, col + i - 1]
                                                .AddComment(
                                                    "Milestone " + aTask.RollupName + " from " + aTask.MSPUID + " '" +
                                                    aTask.Name + string.Format("' finish on {0:yyyy-MM-dd}",
                                                        aTask.Finish.Value),
                                                    $"ontrack on {DateTime.Now.ToString("yyyy-MM-dd")}");
                                            ganttRange[row, col + i - 1].Comment.AutoFit = true;
                                        }
                                        else
                                        {
                                            ganttRange[row, col + i - 1].Comment.Text =
                                                "Milestone " + aTask.RollupName + "from " + aTask.MSPUID + " '" + aTask.Name +
                                                string.Format("' finish on {0:yyyy-MM-dd}", aTask.Finish.Value);

                                            ganttRange[row, col + i - 1].Comment.Author =
                                                $"OnTrackTool on {DateTime.Now.ToString("yyyy-MM-dd")}";

                                            ganttRange[row, col + i - 1].Comment.AutoFit = true;
                                        }
                                }
                           
                            }

                        // recursion if this task is also a summary
                        if (aTask.IsSummary && task.IsRollup) GenerateGanttRowRollup(row, col, aTask);
                              
                    }
                    else
                    {

                    }
                }

                return true;
        }
    }
}
