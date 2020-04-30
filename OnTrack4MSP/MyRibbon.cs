using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using MSProject = Microsoft.Office.Interop.MSProject;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;


// TODO:  Führen Sie diese Schritte aus, um das Element auf dem Menüband (XML) zu aktivieren:

// 1: Kopieren Sie folgenden Codeblock in die ThisAddin-, ThisWorkbook- oder ThisDocument-Klasse.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Erstellen Sie Rückrufmethoden im Abschnitt "Menübandrückrufe" dieser Klasse, um Benutzeraktionen
//    zu behandeln, z.B. das Klicken auf eine Schaltfläche. Hinweis: Wenn Sie dieses Menüband aus dem Menüband-Designer exportiert haben,
//    verschieben Sie den Code aus den Ereignishandlern in die Rückrufmethoden, und ändern Sie den Code für die Verwendung mit dem
//    Programmmodell für die Menübanderweiterung (RibbonX).

// 3. Weisen Sie den Steuerelementtags in der Menüband-XML-Datei Attribute zu, um die entsprechenden Rückrufmethoden im Code anzugeben.  

// Weitere Informationen erhalten Sie in der Menüband-XML-Dokumentation in der Hilfe zu Visual Studio-Tools für Office.


namespace OnTrackMSP
{
    [ComVisible(true)]
    public class MyRibbon : Office.IRibbonExtensibility
    {
        internal const string ConstPropertyName_ProjectId = "OnTrackMSP_ProjectID";
        internal const string ConstPropertyName_LastDBExport = "OnTrackMSP_LastDBExport";
        internal const string ConstPropertyName_DatabaseIdt = "OnTrackMSP_DatabaseId";
        internal const string ConstPropertyName_DatabaseConnectionString = "OnTrackMSP_DatabaseConnectionString";
        internal const string ConstVersion = "01";

        private Office.IRibbonUI ribbon;

        public MyRibbon()
        {

        }

        /// <summary>
        /// write a document property in the project file - overwrite if it iexists
        /// </summary>
        /// <param name="project"></param>
        /// <param name="propertyName"></param>
        /// <param name="value"></param>
        public static void WriteDocumentProperty(MSProject.Project project, string propertyName, string value)
        {
            Microsoft.Office.Core.DocumentProperties properties;
            properties = (Office.DocumentProperties)project.CustomDocumentProperties;

            if (ReadDocumentProperty(project, propertyName) != null)
            {
                properties[propertyName].Delete();
            }

            properties.Add(propertyName, false,
                Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString,
                value);
        }
        
        /// <summary>
        /// return a Office Document Property for a project
        /// </summary>
        /// <param name="project"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        public static string ReadDocumentProperty(MSProject.Project project, string propertyName)
        {
            Office.DocumentProperties properties;
            properties = (Office.DocumentProperties)project.CustomDocumentProperties;

            foreach (Office.DocumentProperty prop in properties)
            {
                if (prop.Name == propertyName)
                {
                    return prop.Value.ToString();
                }
            }
            return null;
        }

        /// <summary>
        /// return the project id from the properties of the current project file
        /// </summary>
        /// <returns></returns>
        public static string GetProjectId(MSProject.Project project = null)
        {
            // Publish the active Project
            if (project == null)
                project = Globals.ThisAddIn.Application.ActiveProject;
            
            string projectId = ReadDocumentProperty(project, ConstPropertyName_ProjectId);
            if (String.IsNullOrEmpty((projectId)))
            {
                projectId = Path.GetFileNameWithoutExtension(project.FullName);
                WriteDocumentProperty(project,ConstPropertyName_ProjectId,projectId);

            }

            return projectId;
        }
        /// <summary>
        /// return the project id from the properties of the current project file
        /// </summary>
        /// <returns></returns>
        public string GetDatabaseConnectionString(MSProject.Project project = null)
        {
            // Publish the active Project
            if (project == null)
                project = Globals.ThisAddIn.Application.ActiveProject;

            string aConnectionString = ReadDocumentProperty(project, ConstPropertyName_DatabaseConnectionString);
            // we have a ConnectionString
            if (!String.IsNullOrEmpty((aConnectionString)))
            {
                // if we donot have a directory -> take directory of current file
                if (String.IsNullOrEmpty(Path.GetDirectoryName(aConnectionString)))
                {
                    var aFilePath = Path.GetDirectoryName(project.FullName);
                    aConnectionString = aFilePath + "\\" + aConnectionString;
                }
            }
            // if we do not have a connection string
            else
            {
                // use the projectId
                string aFileName = GetProjectId(project) + ".db";
                
                if (string.IsNullOrEmpty(aFileName))
                    aFileName = Path.GetFileNameWithoutExtension(project.FullName);
                // set the the property
                WriteDocumentProperty(project, ConstPropertyName_DatabaseConnectionString, aFileName);
                // leave it to the directory
                var aFilePath = Path.GetDirectoryName(project.FullName);
                aConnectionString = aFilePath + "\\" + aFileName;
            }

           

            // return

            return aConnectionString;
        }
        /// <summary>
        /// publish button event
        /// </summary>
        /// <param name="control"></param>
        public void OnPublishButton(Office.IRibbonControl control)
        {
          
            // Publish the active Project
            MSProject.Project myActiveProject = Globals.ThisAddIn.Application.ActiveProject;

            // set database
            var aString = GetDatabaseConnectionString(project: myActiveProject);
            if (DBase.DatabaseConnectionString != aString)
            {
                DBase.DatabaseConnectionString = aString;
                myActiveProject.Application.StatusBar = "Database set to '" + aString + "'";
            }
            
                
            PublishDBase.UpdateAllDbTasks(GetProjectId(),myActiveProject);

            WriteDocumentProperty(myActiveProject,propertyName:ConstPropertyName_LastDBExport, value: DateTime.Now.ToString());

            myActiveProject.Application.StatusBar = "Data from '" + myActiveProject.Name + "' published to Database '" +
                                                    DBase.Id + "'";
        }
        /// <summary>
        /// create excel roadmap event
        /// </summary>
        /// <param name="control"></param>
        public void OnCreateRoadmapButton(Office.IRibbonControl control)
        {
            MSProject.Project myActiveProject = Globals.ThisAddIn.Application.ActiveProject;
            var aFileName = Path.GetFileNameWithoutExtension(myActiveProject.FullName);
            var aFilePath = Path.GetDirectoryName(myActiveProject.FullName);

            if (!String.IsNullOrEmpty(aFilePath)) aFileName = aFilePath + "\\" + aFileName + ".xlsx";
            var aProjectId = GetProjectId(myActiveProject);
            // set database
            var aString = GetDatabaseConnectionString(project: myActiveProject);
            if (DBase.DatabaseConnectionString != aString)
            {
                DBase.DatabaseConnectionString = aString;
                myActiveProject.Application.StatusBar = "Database set to '" + aString + "'";
            }

            PublishExcel.RunExcelRoadmap(aFileName, projectId:aProjectId);
            myActiveProject.Application.StatusBar = "Export Roadmap to " + aFileName + " written";
        }
        /// <summary>
        /// create excel baseline overview event
        /// </summary>
        /// <param name="control"></param>
        public void OnCreateBaselineOverviewButton(Office.IRibbonControl control)
        {
            MSProject.Project myActiveProject = Globals.ThisAddIn.Application.ActiveProject;
            var aFileName = Path.GetFileNameWithoutExtension(myActiveProject.FullName);
            var aFilePath = Path.GetDirectoryName(myActiveProject.FullName);

            if (!String.IsNullOrEmpty(aFilePath)) aFileName = aFilePath + "\\" + aFileName + ".xlsx";
            var aProjectId = GetProjectId(myActiveProject);
            // set database
            var aString = GetDatabaseConnectionString(project: myActiveProject);
            if (DBase.DatabaseConnectionString != aString)
            {
                DBase.DatabaseConnectionString = aString;
                myActiveProject.Application.StatusBar = "Database set to '" + aString + "'";
            }

            PublishExcel.RunExcelBaselineOverview(aFileName, projectId: aProjectId);

            myActiveProject.Application.StatusBar = "Export Baseline Overview to " + aFileName + " written";
        }
        /// <summary>
        /// Update the External Tasks Columns from Database
        /// </summary>
        /// <param name="control"></param>
        public void OnUpdateExternalButton(Office.IRibbonControl control)
        {
            // Publish the active Project
            MSProject.Project myActiveProject = Globals.ThisAddIn.Application.ActiveProject;

            // set database
            var aString = GetDatabaseConnectionString(project: myActiveProject);
            if (DBase.DatabaseConnectionString != aString)
            {
                DBase.DatabaseConnectionString = aString;
                myActiveProject.Application.StatusBar = "Database set to '" + aString + "'";
            }


            MSPUpdateExternalTasks.UpdateExternalTasks(myActiveProject);

            myActiveProject.Application.StatusBar = "Update External Tasks finished";
        }
        #region IRibbonExtensibility-Member

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OnTrackMSP.MyRibbon.xml");
        }

        #endregion

        #region Menübandrückrufe
        //Erstellen Sie hier Rückrufmethoden. Weitere Informationen zum Hinzufügen von Rückrufmethoden finden Sie unter https://go.microsoft.com/fwlink/?LinkID=271226.

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            
        }

        #endregion

        #region Hilfsprogramme

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
