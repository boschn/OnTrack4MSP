using LiteDB;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OnTrackMSP
{
   
    /// <summary>
    ///  Wrapper Class for a project task
    /// </summary>
    public class dbTask
    {
        

        // MS Project Data
        [BsonId]
        public string _id { get; }
        public string ProjectID { get; internal set; }
        public long MSPId { get; set; }
        public long MSPUID { get; internal set; }
        public string GUID { get; internal set;  }
        public string Name { get; set; }
        public DateTime? Start { get; set; }
        public DateTime? Finish { get; set; }
        public ushort Progress { get; set; }
        public bool IsSummary { get; set; }
        public bool IsRollup { get; set; }
        public bool IsMilestone { get; set; }
        public UInt64 OutlineLevel { get; set; }
        public List<string> Resources { get; internal set; }
        public List<string> Predecessors { get; internal set; }
        public List<string> OutlineChildren { get; internal set; }

        public string XtrnCode { get; set; }
        public string XtrnProjectId { get; set; }
        public string Responsible { get; set; }
        public string PlanType { get; set; }
        public string XtrnName { get; set; }
        public string XtrnPredecessor { get; set; }
        public string ProductRelease { get; set; }
        public string RollupName { get; set; }
        public string VariantName { get; set; }
        public string BaselineCode { get; set; }
        public string Delegated { get; set; }
        public string SBSCode { get; set; }
        //custom dates
        public DateTime? XtrnStart { get; set; }
        public DateTime? XtrnFinish { get; set; }

        public Int64 XtrnProgress { get; set; }

        // custom flags
        public bool IsRoadmap { get; set; }
        public bool IsXtrnDriven { get; set; }
        public bool IsGovernance { get; set; }
        public bool IsBaseline { get; set; }
        public bool IsLocoNeeded { get; set; }
        public bool IsPendTime { get; set; }
        public bool IsRollUpSummary { get; set; }
        public DateTime? ActualStart { get; set; }
        public DateTime? ActualFinish { get; set; }
        
        public long UpdateStamp { get; set; }
        public DateTime? XtrnUpdated { get; set; }

        private dbTask()
        {
            Debug.Write("upps");
        }

        /// <summary>
        /// convert a projectId and a MSPUID to the universal ID)
        /// </summary>
        /// <param name="projectId"></param>
        /// <param name="UID"></param>
        /// <returns></returns>
        public static string Convert2UniqueKey(string projectId, long UID)
        {
            return projectId + ":" + UID;
        }
        public static string Convert2UniqueKey(string projectId, string UID)
        {
            return projectId + ":" + UID;
        }
        public static long? RetrieveUid(string uniqueKey)
        {
            if (String.IsNullOrEmpty(uniqueKey)) return null;
            var theParts = uniqueKey.Split(':');
            if ((theParts != null) && (theParts.IsFixedSize))
            {
                return PublishDBase.Convert2Long(theParts[1]);
            }

            return null;
        }

        /// <summary>
        /// CTOR for BSON
        /// </summary>
        /// <param name="projectID"></param>
        /// <param name="mspuid"></param>
        [BsonCtor]
        public dbTask(string projectID, Int64 mspuid)
        {
            this._id = Convert2UniqueKey(projectID,mspuid);
            this.ProjectID = projectID;
            this.MSPUID = mspuid;
            this.GUID = RandomString(20,false);
            this.Predecessors = new List<string>();
            this.OutlineChildren = new List<string>();
            this.Resources = new List<string>();
            // this.UpdateStamp = DateTime.Now;
            this.ActualFinish = new DateTime?();
            this.ActualStart = new DateTime?();
            this.Finish = new DateTime?();
            this.Start = new DateTime?();
            this.XtrnFinish = new DateTime?();
            this.XtrnStart = new DateTime?();
            this.XtrnUpdated = new DateTime?();
        }
        // Generate a random string with a given size    
        private static string RandomString(int size, bool lowerCase)
        {
            StringBuilder builder = new StringBuilder();
            Random random = new Random();
            char ch;
            for (int i = 0; i < size; i++)
            {
                ch = Convert.ToChar(Convert.ToInt32(Math.Floor(26 * random.NextDouble() + 65)));
                builder.Append(ch);
            }
            if (lowerCase)
                return builder.ToString().ToLower();
            return builder.ToString();
        }
    }
}
