using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using MSProject = Microsoft.Office.Interop.MSProject;
using Office = Microsoft.Office.Core;

namespace OnTrackMSP
{
    //
    // publish static class 
    //

    /// <summary>
    /// static publish class
    /// </summary>
    static class PublishDBase
    {
        
        /// <summary>
        /// Convert to Bool value - Hack -
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool Convert2Bool (String value)
        {
            try
            {
                // Hack
                if (string.Compare(value, "ja", true) == 0) return (true);
                if (string.Compare(value, "yes", true) == 0) return (true);
                if (string.Compare(value, "true", true) == 0) return (true);
                if (string.Compare(value, "1", true) == 0) return (true);
                // return (Convert.ToBoolean(value)); -> throws exception if not false
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine(ex);
                return false;
            }
            
            return false;
        }

        /// <summary>
        /// Convert a value to long
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static long Convert2Long(String value)
        {

            try
            {
            	return Convert.ToInt64(value);
            }
            catch (System.Exception ex)
            {

                return 0;
            }
        }

        public static DateTime? Convert2DateTime (object value)
        {
          
            try
            {
                string aValue = Convert.ToString(value);
                if (String.Compare(aValue, "NV", StringComparison.CurrentCultureIgnoreCase) == 0) return null;
            	return Convert.ToDateTime(value);
            }
            catch (System.Exception ex)
            {
                return null;
            }
        }

        /// <summary>
        /// publish a ms project schedule to the database
        /// </summary>
        /// <param name="project"></param>
        public static void UpdateAllDbTasks(string projectId, MSProject.Project project)
        {
            var aStamp = DateTime.Now.Ticks;

            // Iterating over tasks in active project
            foreach (MSProject.Task aMSPTask in project.Tasks)
            {
                // get a New dbTask 
                if (aMSPTask != null)
                {
                    var aTaskDb = DBase.GetTask(dbTask.Convert2UniqueKey(projectId: projectId, aMSPTask.UniqueID));
                    if (aTaskDb == null)
                    {
                        aTaskDb = new dbTask(projectId, aMSPTask.UniqueID);
                        aTaskDb.UpdateStamp = aStamp;
                        DBase.InsertTask(aTaskDb);
                    }

                    // donot use not active 
                    if (!aMSPTask.Active)
                    {
                        DBase.DeleteTask(aTaskDb);
                    }
                    else
                    {
                        UpdateTaskDb(aTaskDb, aMSPTask, timestamp: aStamp);
                    }
                }

            }

            // run for deleted MSP tasks 
            //
            foreach (var aTaskDb in DBase.GetTasks(projectId))
            {
                if (aTaskDb.UpdateStamp != aStamp) 
                    DBase.DeleteTask(aTaskDb);
            }

            // messagebox
            MessageBox.Show(icon: MessageBoxIcon.Information, caption: "Publish to Database done", text: DBase.GetTasks().Count() + @" records updated", buttons: MessageBoxButtons.OK);
        }

        /// <summary>
        /// update a single task from the database according to a leading msptask
        /// </summary>
        /// <param name="taskdb"></param>
        /// <param name="msptask"></param>
        /// <returns></returns>
        internal static bool UpdateTaskDb(dbTask taskdb, MSProject.Task msptask, long? timestamp = null)
        {
            try
            {
                taskdb.Name = msptask.Name;
                taskdb.MSPId = msptask.ID;
                taskdb.MSPUID = msptask.UniqueID;
                taskdb.GUID = msptask.Guid;
                taskdb.Start = Convert2DateTime(msptask.Start);

                taskdb.ActualStart = Convert2DateTime(msptask.ActualStart);
                taskdb.Finish = Convert2DateTime(msptask.Finish);
                taskdb.ActualFinish = Convert2DateTime(msptask.ActualFinish);
                taskdb.IsSummary = msptask.Summary;
                taskdb.IsMilestone = msptask.Milestone;
                taskdb.OutlineLevel = Convert.ToUInt64(msptask.OutlineLevel);
                taskdb.IsRollup = msptask.Rollup;

                // resourceNames
                taskdb.Resources.Clear();
                if (!String.IsNullOrWhiteSpace(msptask.ResourceNames))
                    foreach (string aValue in msptask.ResourceNames.Split(new char[] { ',', ';' }))
                    {
                        if (!taskdb.Resources.Contains(aValue.ToUpper()))
                            taskdb.Resources.Add(aValue.ToUpper());
                    }


                // predecessors lazy
                taskdb.Predecessors.Clear(); // this is easier and lazy
                if (!String.IsNullOrWhiteSpace(msptask.UniqueIDPredecessors))
                    foreach (string aValue in msptask.UniqueIDPredecessors.Split(new char[] { ',', ';' }))
                    {
                        if (!String.IsNullOrWhiteSpace(aValue))
                        {
                            string pattern = @"^\d{4}";
                            System.Text.RegularExpressions.Regex r =
                                new System.Text.RegularExpressions.Regex(pattern);
                            var groups = r.Matches(aValue);
                            if (groups.Count > 0)
                            {
                                var anUid = Convert2Long(groups[0].Value);
                                var aKey = dbTask.Convert2UniqueKey(taskdb.ProjectID, anUid);
                                // add the uid of the predecessor - skip the rest of the contraint
                                if (!taskdb.Predecessors.Contains(aKey)) taskdb.Predecessors.Add(aKey);
                            }

                        }
                    }

                // outline children lazy
                taskdb.OutlineChildren.Clear();
                if (msptask.OutlineChildren != null)
                    foreach (MSProject.Task aSubTaskMSP in msptask.OutlineChildren)
                    {
                        var aKey = dbTask.Convert2UniqueKey(taskdb.ProjectID, aSubTaskMSP.UniqueID);
                        if (!taskdb.OutlineChildren.Contains(aKey))
                            taskdb.OutlineChildren.Add(aKey);
                    }


                // custom fields mapping

                taskdb.IsLocoNeeded = Convert2Bool(msptask.GetField(MSProject.PjField.pjTaskFlag1));
                taskdb.IsPendTime = Convert2Bool(msptask.GetField(MSProject.PjField.pjTaskFlag2));
                taskdb.IsRoadmap = Convert2Bool(msptask.GetField(MSProject.PjField.pjTaskFlag4));
                taskdb.IsXtrnDriven = Convert2Bool(msptask.GetField(MSProject.PjField.pjTaskFlag5));
                taskdb.IsGovernance = Convert2Bool(msptask.GetField(MSProject.PjField.pjTaskFlag6));
                taskdb.IsBaseline = Convert2Bool(msptask.GetField(MSProject.PjField.pjTaskFlag7));

                taskdb.XtrnCode = msptask.GetField(MSProject.PjField.pjTaskText11);
                taskdb.XtrnProjectId = msptask.GetField(MSProject.PjField.pjTaskText10);
                taskdb.Delegated = msptask.GetField(MSProject.PjField.pjTaskText2);
                taskdb.Responsible = msptask.GetField(MSProject.PjField.pjTaskText3).ToUpper();
                taskdb.PlanType = msptask.GetField(MSProject.PjField.pjTaskText4).ToUpper();
                taskdb.XtrnName = msptask.GetField(MSProject.PjField.pjTaskText5);
                taskdb.XtrnPredecessor = msptask.GetField(MSProject.PjField.pjTaskText6);
                taskdb.ProductRelease = msptask.GetField(MSProject.PjField.pjTaskText7).ToUpper();
                taskdb.RollupName = msptask.GetField(MSProject.PjField.pjTaskText8);

                taskdb.VariantName = msptask.GetField(MSProject.PjField.pjTaskOutlineCode1).ToUpper();
                taskdb.BaselineCode = msptask.GetField(MSProject.PjField.pjTaskOutlineCode2).ToUpper();
                taskdb.SBSCode = msptask.GetField(MSProject.PjField.pjTaskOutlineCode3).ToUpper();

                taskdb.XtrnPredecessor = msptask.GetField(MSProject.PjField.pjTaskText19);
                taskdb.XtrnProgress = Convert2Long(msptask.GetField(MSProject.PjField.pjTaskNumber3));
                taskdb.XtrnStart = Convert2DateTime(msptask.GetField(MSProject.PjField.pjTaskStart1));
                taskdb.XtrnFinish = Convert2DateTime(msptask.GetField(MSProject.PjField.pjTaskFinish1));
                taskdb.XtrnUpdated = Convert2DateTime(msptask.GetField(MSProject.PjField.pjTaskDate10));
                // update
                if (!timestamp.HasValue) timestamp =  DateTime.Now.Ticks;
                
                taskdb.UpdateStamp = timestamp.Value;
                DBase.UpdateTask(taskdb);

            }
            catch (System.Exception ex)
            {
                Debug.WriteLine(ex);
                return false;
            }

            return true;

        }
        /// <summary>
        /// returns true if a string is a date
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        static bool CheckDate(object date)
        {
            DateTime Temp;

            if (date == null) return false;

            if (DateTime.TryParse(date.ToString(), out Temp) == true)
                return true;
            else
                return false;
        }
    }
  
}
