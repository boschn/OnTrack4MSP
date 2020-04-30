using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace OnTrackMSP
{

    /// <summary>
    /// class to publish via EPPL
    /// </summary>
    class PublishExcel
    {
        private static IEnumerable<dbTask> ourTasks ;

        /// <summary>
        /// publish Database to an Excel
        /// </summary>
        public static void RunExcelRoadmap (string filename, string projectId = null)
        {
            ourTasks = DBase.GetTasks();
            // all product releases
            foreach (var aProductRelease in DBase.GetProductReleases())
            {
                // create list of all tasks
                var aResultList = DBase.GetTasks(projectId)
                    .Where(x => (x.IsRoadmap == true) && (x.ProductRelease==aProductRelease))
                    .OrderBy(x => x.SBSCode)
                    .ToList();
                // write
                WriteExcelRoadmap(aProductRelease,aResultList, filename);
            }
            
        }
        /// <summary>
        /// publish Database to an Excel
        /// </summary>
        public static void WriteExcelRoadmap(string productrelease, List<dbTask> tasklist, string filename)
        {
            //Open the workbook (or create it if it doesn't exist)
            var aFile = new FileInfo(filename);
            
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            try
            {
                using (var p = new ExcelPackage(aFile))
                {

                    //Create a worksheet "ProductRoadMap " + productrelease

                    var ws = p.Workbook.Worksheets.FirstOrDefault(s => String.Compare(s.Name, "ProductRoadMap " + productrelease, StringComparison.CurrentCultureIgnoreCase) == 0);
                    if (ws == null) ws = p.Workbook.Worksheets.Add("ProductRoadMap " + productrelease);

                    var aGantt = new PublishExcelGantt("ProductRoadMap " + productrelease, ws.SelectedRange, tasklist);
                    aGantt.Generate(PublishExcelGantt.RoadmapTableDef);

                    //Save and close the package.
                    p.Save();
                }
            }
            catch (System.IO.IOException ex) 
            {
                MessageBox.Show(icon: MessageBoxIcon.Error, caption: "Error", text: ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.Data, buttons: MessageBoxButtons.OK);
            }
            catch (Exception ex) when (!Env.Debugging)
            {
                MessageBox.Show(icon: MessageBoxIcon.Error, caption: "Error", text: ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.Data, buttons: MessageBoxButtons.OK);
            }
        
           
        }
        /// <summary>
        /// publish Database to an Excel
        /// </summary>
        public static void RunExcelBaselineOverview(string filename, string projectId = null)
        {

            ourTasks = DBase.GetTasks();
            List<dbTask> aResultList = new List<dbTask>();

            // create a list ordered by baseline code, variantname, finish date ascending
            foreach (var aBaselineCode in DBase.GetBaselineCodes())
            {
                // create list of tasks from baseline to baseline ordered by Variantnames
                var aBaselineList = DBase.GetTasks(projectId)
                    .Where(x => (x.IsBaseline == true) && (x.BaselineCode == aBaselineCode))
                    .OrderBy(x => x.VariantName)
                    .ToList();

                foreach (var aVariantname in DBase.GetVariantNames())
                {
                    var aNameList = 
                            from task in aBaselineList
                                  where task.VariantName == aVariantname
                                  orderby task.Finish ascending 
                                  select task;

                    aResultList.AddRange(aNameList);
                }
               
            }

            // write
            WriteExcelBaselineOverview(aResultList, filename);

            
        }
        /// <summary>
        /// publish Database to an Excel
        /// </summary>
        public static void WriteExcelBaselineOverview(List<dbTask> tasklist, string filename)
        {
            //Open the workbook (or create it if it doesn't exist)
            var aFile = new FileInfo(filename);
          
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var p = new ExcelPackage(aFile))
                {

                    //Create a worksheet "BaselineOverview " 

                    var ws = p.Workbook.Worksheets.FirstOrDefault(s => String.Compare(s.Name, "BaselineOverview", StringComparison.CurrentCultureIgnoreCase) == 0);
                    if (ws == null) ws = p.Workbook.Worksheets.Add("BaselineOverview");

                    var aGantt = new PublishExcelGantt("BaselineOverview", ws.SelectedRange, tasklist);
                    aGantt.Generate(PublishExcelGantt.BaselineOverViewTableDef);
                    //Save and close the package.
                    p.Save();
                }
            } 
            catch(System.IO.IOException ex)
            {
                MessageBox.Show(icon: MessageBoxIcon.Error, caption: "Error", text: ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.Data, buttons: MessageBoxButtons.OK);
            }
            catch (Exception ex) when (!Env.Debugging)
            {
                MessageBox.Show(icon: MessageBoxIcon.Error, caption: "Error", text: ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.Data, buttons: MessageBoxButtons.OK);
            }
        }

    }
}
