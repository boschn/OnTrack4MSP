using LiteDB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;

namespace OnTrackMSP
{
    static class DBase

    {

        /// <summary>
        /// return the DatabaseConnectionString
        /// </summary>
        internal static string DatabaseConnectionString
        {
            get => String.IsNullOrEmpty(ourDatabaseConnectionString) == true ? ConstDatabaseDefaultConnectionString : ourDatabaseConnectionString;
            set
            {
                // reset database
                if (ourDatabaseConnectionString != value)
                {
                    ourDatabaseConnectionString = value;
                    ourTasksDB = null;
                    if (ourLiteDB != null) {ourLiteDB.Dispose(); ourLiteDB = null;}
                    Id = Path.GetFileNameWithoutExtension(value);
                }

            } 
        }
        internal static string Id { get; set; }
        static private LiteDatabase ourLiteDB;
        static private LiteCollection<dbTask> ourTasksDB;
        private static string ourDatabaseConnectionString;

        // const names
        private const string ConstTasksObjectName = "tasks";
        private const string ConstDatabaseDefaultConnectionString = @"myLiteData.db";



        /// <summary>
        /// return database
        /// </summary>
        /// <returns></returns>
        internal static LiteDatabase GetDb ()
        {
            if (DBase.ourLiteDB == null)
            {
                ourLiteDB = new LiteDatabase(connectionString: ourDatabaseConnectionString) ;
                
            }
            return ourLiteDB;
        }

        /// <summary>
        /// returns the LiteCollection of dbTask or creates it
        /// </summary>
        /// <returns></returns>
        private static LiteCollection<dbTask> GetAllTasks ()
        {
            if (ourTasksDB == null)
            {
                ourTasksDB = (LiteCollection<dbTask>)GetDb().GetCollection<dbTask>(name: ConstTasksObjectName);
                // ourTasksDB.EnsureIndex(x => x.GUID); -> doesnt work
            }

            return ourTasksDB;
        }

        /// <summary>
        /// insert a task in the database
        /// </summary>
        /// <param name="task"></param>
        /// <returns>true if successfull</returns>
        internal static bool InsertTask(dbTask task)
        {
            try
            {
                GetAllTasks().Insert(task);
                return true;
            }
            catch (Exception e)
            {
                return false;
            }

        }

        /// <summary>
        /// delete a task from the database
        /// </summary>
        /// <param name="task"></param>
        /// <returns>true if successfull</returns>
        internal static  bool DeleteTask(dbTask task)
        {
            try
            {
                GetAllTasks().Delete(task._id);
                return true;
            }
            catch (Exception e)
            {
                return false;
            }

        }

        /// <summary>
        /// Update a task to the database
        /// </summary>
        /// <param name="task"></param>
        /// <returns>true if succesfull</returns>
        internal static bool UpdateTask(dbTask task)
        {
            try
            {
                GetAllTasks().Update(task);
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }
        /// <summary>
        /// load all tasks as collection or load the collection
        /// </summary>
        /// <returns></returns>
        internal static IList<dbTask> GetTasks(string projectId = null)
        {
            
            if (projectId != null)
                if (GetAllTasks().Count() != 0)
                    return GetAllTasks().Query().Where(x => x.ProjectID == projectId).ToList();
                else return new List<dbTask>();
            else if (GetAllTasks().Count() != 0)
            {
                
                return GetAllTasks().FindAll().ToList();
            }
               
            else return new List<dbTask>();
        }

        /// <summary>
        /// return a task object by id (unique-key)
        /// </summary>
        /// <param name="id"></param>
        /// <returns>dbTask</returns>
        internal static dbTask GetTask(string id)
        {
            if (ourTasksDB == null) GetAllTasks();
            return ourTasksDB.FindById(id);
        }
        /// <summary>
        /// load all tasks as collection or load the collection
        /// </summary>
        /// <returns></returns>
        internal static SortedSet<string> GetProductReleases(string projectId = null)
        {
            List<String> results;
            var aSet = new SortedSet<string>();

            // Use LINQ to query documents 
            if (projectId == null)
                results = GetTasks()
                .Where(x => !String.IsNullOrEmpty(x.ProductRelease))
                .Select(x => x.ProductRelease)
                .ToList();
            else
                results = GetTasks()
                    .Where(x => !String.IsNullOrEmpty(x.ProductRelease) && x.ProjectID == projectId)
                    .Select(x => x.ProductRelease)
                    .ToList();
            

            foreach (var aValue in results)
            {
                aSet.Add(aValue);
            }

            return aSet;
        }
        /// <summary>
        /// load all tasks as collection or load the collection
        /// </summary>
        /// <returns></returns>
        internal static SortedSet<string> GetBaselineCodes(string projectId = null)
        {
            List<String> results;
            var aSet = new SortedSet<string>();
            // Use LINQ to query documents 
            if (projectId == null)
                results = GetTasks()
                .Where(x => !String.IsNullOrEmpty(x.BaselineCode))
                .Select(x => x.BaselineCode)
                .ToList();
            else
                results = GetTasks()
                    .Where(x => !String.IsNullOrEmpty(x.BaselineCode) && x.ProjectID == projectId)
                    .Select(x => x.BaselineCode)
                    .ToList();
            foreach (var aValue in results)
            {
                aSet.Add(aValue);
            }

            return aSet;
        }
        /// <summary>
        /// load all tasks as collection or load the collection
        /// </summary>
        /// <returns></returns>
        internal static SortedSet<string> GetVariantNames(string projectId = null)
        {
            List<String> results;
            var aSet = new SortedSet<string>();
            // Use LINQ to query documents 
            if (projectId == null) results = GetTasks()
                .Where(x => !String.IsNullOrEmpty(x.VariantName))
                .Select(x => x.VariantName)
                .ToList();
            else results = GetTasks()
                .Where(x => !String.IsNullOrEmpty(x.VariantName) && x.ProjectID == projectId )
                .Select(x => x.VariantName)
                .ToList();

            
            foreach (var aValue in results)
            {
                aSet.Add(aValue);
            }

            return aSet;
        }
        /// <summary>
        /// load all tasks as collection or load the collection
        /// </summary>
        /// <returns></returns>
        internal static SortedSet<string> GetProjectIds ()
        {
            List<String> results;
            var aSet = new SortedSet<string>();

            // Use LINQ to query documents 
            results = GetTasks()
                .Where(x => !String.IsNullOrEmpty(x.VariantName))
                .Select(x => x.ProjectID)
                .ToList();

            foreach (var aValue in results)
            {
                aSet.Add(aValue);
            }

            return aSet;
        }

    }
}
