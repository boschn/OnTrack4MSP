using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace OnTrackMSP
{
    class MSPUpdateExternalTasks


    {

        internal static string RemoveWhitespace(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";

            return new string(input.ToCharArray()
                .Where(c => !Char.IsWhiteSpace(c))
                .ToArray());
        }
        internal static bool UpdateExternalTasks( MSProject.Project project, string defaultExternalProjectId = null)
        {
            var aStamp = DateTime.Now;

            // Iterating over tasks in active project
            foreach (MSProject.Task aMSPTask in project.Tasks)
            {
                if (aMSPTask != null)
                {
                    // project id of current project
                    var aProjectId = MyRibbon.GetProjectId(project);
                    // see if we have external data
                    var IsXtrnDriven = PublishDBase.Convert2Bool(aMSPTask.GetField(MSProject.PjField.pjTaskFlag5));
                    var XtrnCode = aMSPTask.GetField(MSProject.PjField.pjTaskText11);
                    var XtrnProjectId = aMSPTask.GetField(MSProject.PjField.pjTaskText10);
                    if (String.IsNullOrEmpty(XtrnProjectId)) XtrnProjectId = defaultExternalProjectId;

                    // get a current task
                    var aTaskDb = DBase.GetTask(dbTask.Convert2UniqueKey(projectId: aProjectId, aMSPTask.UniqueID));

                    // donot use not active 
                    if ((!aMSPTask.Active) && (aTaskDb != null))
                    {
                        DBase.DeleteTask(aTaskDb);
                    }
                    else
                    {
                        try
                        {
                            // get the external task

                            if (!String.IsNullOrEmpty(XtrnCode) && !String.IsNullOrEmpty(XtrnProjectId))
                            {
                                // if the XtrnCode contains more than one reference
                                //

                                var theExtUids = XtrnCode.Split(',');
                                uint exprNo = 0;

                                // init the fields
                                aMSPTask.Finish1 = "NV";
                                aMSPTask.Start1 = "NV";


                                foreach (var aValue in theExtUids)
                                {
                                    string anExpression = "";
                                    string anExtUid = "";
                                    exprNo++;

                                    // remove Subexpression
                                    if (aValue.Contains(':'))
                                    {
                                        var theSubExp = aValue.Split(':');
                                        anExtUid = RemoveWhitespace(theSubExp[0]);
                                        anExpression = RemoveWhitespace(theSubExp[1]);

                                    }
                                    else anExtUid = RemoveWhitespace(aValue);

                                    // get the external Tasks
                                    var anExternalTaskDb =
                                        DBase.GetTask(dbTask.Convert2UniqueKey(projectId: XtrnProjectId, anExtUid));
                                    if (anExternalTaskDb == null)
                                    {
                                        var aMessage = aMSPTask.Text18;
                                        if (!String.IsNullOrEmpty(aMessage)) aMessage += ", ";
                                        aMSPTask.SetField(MSProject.PjField.pjTaskText18, aMessage +
                                                                                          "External Task " + anExtUid +
                                                                                          " not found " + DateTime.Now);
                                        exprNo--; // reduce again
                                    }
                                    else
                                    {
                                        // sum the name according to the external task
                                        if (exprNo == 1)
                                            aMSPTask.SetField(MSProject.PjField.pjTaskText5, anExternalTaskDb.Name);
                                        else
                                        {
                                            var aString = aMSPTask.GetField((MSProject.PjField.pjTaskText5));
                                            aString += "," + anExternalTaskDb.Name;
                                            aMSPTask.SetField(MSProject.PjField.pjTaskText5, aString);
                                        }

                                        // sum External Predessors
                                        {
                                            SortedSet<string> thePredecessors = new SortedSet<string>();

                                            foreach (var aPredUID in anExternalTaskDb.Predecessors)
                                            {
                                                thePredecessors.Add(aPredUID);
                                            }

                                            // load the last results 
                                            if (exprNo > 1)
                                            {
                                                var aString2 = aMSPTask.GetField((MSProject.PjField.pjTaskText6));
                                                foreach (var aPredUic in aString2.Split(','))
                                                    if (!String.IsNullOrEmpty(aPredUic))
                                                        thePredecessors.Add(aPredUic);
                                            }

                                            string aStringValue = "";
                                            foreach (var aPredUid in thePredecessors)
                                            {
                                                if (!string.IsNullOrEmpty(aStringValue)) aStringValue += ",";
                                                aStringValue += aPredUid;
                                            }

                                            // update
                                            aMSPTask.SetField(MSProject.PjField.pjTaskText6, aStringValue);
                                        }
                                        // sum up all External resources
                                        {
                                            SortedSet<string> theResources = new SortedSet<string>();

                                            foreach (var aResourcename in anExternalTaskDb.Resources)
                                            {
                                                theResources.Add(aResourcename);
                                            }

                                            // load the last results 
                                            if (exprNo > 1)
                                            {
                                                var aString2 = aMSPTask.GetField((MSProject.PjField.pjTaskText19));
                                                foreach (var aResourcename in aString2.Split(','))
                                                    if (!String.IsNullOrEmpty(aResourcename))
                                                        theResources.Add(aResourcename);

                                            }

                                            string aStringValue = "";
                                            foreach (var aResourcename in theResources)
                                            {
                                                if (!string.IsNullOrEmpty(aStringValue)) aStringValue += ",";
                                                aStringValue += aResourcename;
                                            }

                                            // update
                                            aMSPTask.SetField(MSProject.PjField.pjTaskText19, aStringValue);


                                        }

                                        // see the progress
                                        /* to be implemented
                                        if (anExternalTaskDb.Progress )
                                        aMSPTask.SetField(MSProject.PjField.pjTaskNumber3, anExternalTaskDb.Progress.ToString());
                                        */

                                        // get the start date
                                        // choose the minimum
                                        if (anExpression.ToUpper().Contains("START") ||
                                            String.IsNullOrEmpty(anExpression))
                                        {
                                            if (anExternalTaskDb.Start.HasValue)
                                            {
                                                if ((exprNo == 1) || (String.Compare(aMSPTask.Start1, "NV") == 0)
                                                                  || (DateTime.Compare(aMSPTask.Start1,
                                                                      anExternalTaskDb.Start.Value) > 0))
                                                {
                                                    aMSPTask.Start1 = anExternalTaskDb.Start.Value;
                                                    aMSPTask.SetField(MSProject.PjField.pjTaskStart1,
                                                        anExternalTaskDb.Start.Value.ToString());
                                                }
                                            }
                                        }

                                        // get the finish date
                                        // choose the maximum
                                        if (anExpression.ToUpper().Contains("FINISH") ||
                                            String.IsNullOrEmpty(anExpression))
                                        {
                                            if (anExternalTaskDb.Finish.HasValue)
                                            {
                                                if ((exprNo == 1) || (String.Compare(aMSPTask.Finish1, "NV") == 0)
                                                                  || (DateTime.Compare(aMSPTask.Finish1,
                                                                      anExternalTaskDb.Finish.Value) < 0))
                                                {
                                                    aMSPTask.Finish1 = anExternalTaskDb.Finish.Value;
                                                    aMSPTask.SetField(MSProject.PjField.pjTaskFinish1,
                                                        anExternalTaskDb.Finish.Value.ToString());
                                                }

                                            }
                                        }




                                    }
                                }

                                // run the task if it is planned manual AND external driven
                                //
                                if (IsXtrnDriven && aMSPTask.Manual)
                                {
                                    aMSPTask.Start = aMSPTask.Start1; // min of aTaskDb.Start.Value;
                                    aMSPTask.Finish = aMSPTask.Finish1; // max of aTaskDb.Finish.Value;
                                }

                                // XTernUpdated
                                aMSPTask.Date10 = aStamp;
                                aMSPTask.SetField(MSProject.PjField.pjTaskDate10, aStamp.ToString());
                                {
                                    var aMessage = aMSPTask.Text18;
                                    if (!String.IsNullOrEmpty(aMessage)) aMessage += ", ";
                                    aMSPTask.SetField(MSProject.PjField.pjTaskText18, aMessage +
                                                                                      "UpdateStamp from external Task " +
                                                                                      DateTime.Now);
                                }

                            }
                        }
                        catch (System.Exception ex)
                        {
                            Debug.WriteLine(ex);
                        }
                    }
                }

            }

            MessageBox.Show(DBase.GetTasks().Count() + @" records updated");

            return true;
        }
    }
}
