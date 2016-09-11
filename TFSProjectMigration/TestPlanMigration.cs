using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.Server;
using log4net.Config;
using log4net;
using System.Windows.Controls;
using System.IO;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Framework.Common;

namespace TFSProjectMigration
{
    public class TestPlanMigration
    {
        ITestManagementTeamProject sourceproj;
        ITestManagementTeamProject destinationproj;
        public Hashtable workItemMap;
        ProgressBar progressBar;
        public WorkItemTypeCollection workItemTypes;
        String projectName;
        private static readonly ILog logger = LogManager.GetLogger(typeof(TFSWorkItemMigrationUI));
        public WorkItemStore store;
        private IIdentityManagementService ims;
        private TeamFoundationIdentity UserID;

        public TestPlanMigration(TfsTeamProjectCollection sourceTfs, TfsTeamProjectCollection destinationTfs, string sourceProject, string destinationProject, Hashtable workItemMap, ProgressBar progressBar)
        {
            this.sourceproj = GetProject(sourceTfs, sourceProject);
            this.destinationproj = GetProject(destinationTfs, destinationProject);
            this.workItemMap = workItemMap;
            this.progressBar = progressBar;
            projectName = sourceProject;
            store = (WorkItemStore)destinationTfs.GetService(typeof(WorkItemStore));
            workItemTypes = store.Projects[destinationProject].WorkItemTypes;

            ims = (IIdentityManagementService)sourceTfs.GetService(typeof(IIdentityManagementService));
            UserID = ims.ReadIdentity(IdentitySearchFactor.MailAddress, "hema.kumar@accenture.com", MembershipQuery.Direct, ReadIdentityOptions.None);
        }

        private ITestManagementTeamProject GetProject(TfsTeamProjectCollection tfs, string project)
        {
            
            ITestManagementService tms = tfs.GetService<ITestManagementService>();

            return tms.GetTeamProject(project);
        }
        public void CopyTestPlans()
        {
            int i = 1;
            int planCount= sourceproj.TestPlans.Query("Select * From TestPlan").Count;
            //delete Test Plans if any existing test plans.
            //foreach (ITestPlan destinationplan in destinationproj.TestPlans.Query("Select * From TestPlan"))
            //{

            //    System.Diagnostics.Debug.WriteLine("Deleting Plan - {0} : {1}", destinationplan.Id, destinationplan.Name);

            //    destinationplan.Delete(DeleteAction.ForceDeletion); ;

            //}
           
            foreach (ITestPlan sourceplan in sourceproj.TestPlans.Query("Select * From TestPlan"))
            {
                if (!sourceplan.Name.Equals("CLEF Testing"))
                {
                    continue;
                }
                System.Diagnostics.Debug.WriteLine("Plan - {0} : {1}", sourceplan.Id, sourceplan.Name);

                ITestPlan destinationplan = destinationproj.TestPlans.Create();

                destinationplan.Name = "CLEF Testing V1.0";
                destinationplan.Description = sourceplan.Description;
                destinationplan.StartDate = sourceplan.StartDate;
                destinationplan.EndDate = sourceplan.EndDate;
                destinationplan.State = sourceplan.State;
                destinationplan.Owner = UserID;
                destinationplan.Save();

                //drill down to root test suites.
                if (sourceplan.RootSuite != null && sourceplan.RootSuite.Entries.Count > 0)
                {
                    CopyTestSuites(sourceplan, destinationplan);
                }

                destinationplan.Save();

                progressBar.Dispatcher.BeginInvoke(new Action(delegate()
                {
                    float progress = (float)i / (float) planCount;

                    progressBar.Value = ((float)i / (float) planCount) * 100;
                }));
                i++;
            }

        }

        //Copy all Test suites from source plan to destination plan.
        private void CopyTestSuites(ITestPlan sourceplan, ITestPlan destinationplan)
        {
            ITestSuiteEntryCollection suites = sourceplan.RootSuite.Entries;
            CopyTestCases(sourceplan.RootSuite, destinationplan.RootSuite);

            foreach (ITestSuiteEntry suite_entry in suites)
            {
                IStaticTestSuite suite = suite_entry.TestSuite as IStaticTestSuite;               
                if (suite != null)
                {
                    IStaticTestSuite newSuite = destinationproj.TestSuites.CreateStatic();
                    newSuite.Title = suite.Title;
                    //foreach (var tpoint in newSuite.TestSuiteEntry.PointAssignments)
                    //{
                    //    tpoint.AssignedTo = UserID;
                    //}
                    destinationplan.RootSuite.Entries.Add(newSuite);
                    destinationplan.Save();
                    CopyTestCases(suite, newSuite);
                    if (suite.Entries.Count > 0)
                        CopySubTestSuites(suite, newSuite);
                }
            }

        }

        //Drill down and Copy all subTest suites from source root test suite to destination plan's root test suites.
        private void CopySubTestSuites(IStaticTestSuite parentsourceSuite, IStaticTestSuite parentdestinationSuite)
        {
            ITestSuiteEntryCollection suitcollection = parentsourceSuite.Entries;
            foreach (ITestSuiteEntry suite_entry in suitcollection)
            {
                IStaticTestSuite suite = suite_entry.TestSuite as IStaticTestSuite;
                if (suite != null)
                {
                    IStaticTestSuite subSuite = destinationproj.TestSuites.CreateStatic();
                    subSuite.Title = suite.Title;
                    parentdestinationSuite.Entries.Add(subSuite);

                    CopyTestCases(suite, subSuite);

                    if (suite.Entries.Count > 0)
                        CopySubTestSuites(suite, subSuite);

                }
            }


        }

        //Copy all subTest suites from source root test suite to destination plan's root test suites.
        private void CopyTestCases(IStaticTestSuite sourcesuite, IStaticTestSuite destinationsuite)
        {

            ITestSuiteEntryCollection suiteentrys = sourcesuite.TestCases;

            foreach (ITestSuiteEntry testcase in suiteentrys)
            {
                try
                {   //check whether testcase exists in new work items(closed work items may not be created again).
                    WorkItem newWorkItem = null;
                    Hashtable fieldMap = ListToTable((List<object>)workItemMap[testcase.TestCase.WorkItem.Type.Name]);
                    newWorkItem = new WorkItem(workItemTypes["Test Case"]);
                    foreach (Field field in testcase.TestCase.WorkItem.Fields)
                    {
                        if (field.Name.Contains("ID") || field.Name.Contains("Reason"))
                        {
                            continue;
                        }
                        if (field.Name == "Assigned To" || field.Name == "Activated By")
                        {
                            testcase.TestCase.WorkItem.Open();
                            testcase.TestCase.WorkItem.Fields[field.Name].Value = "hema.kumar@accenture.com";
                        }
                        if (newWorkItem.Fields.Contains(field.Name) && newWorkItem.Fields[field.Name].IsEditable)
                        {
                            newWorkItem.Fields[field.Name].Value = field.Value;
                            if (field.Name == "Iteration Path" || field.Name == "Area Path" || field.Name == "Node Name" || field.Name == "Team Project")
                            {
                                try
                                {
                                    string itPath = (string)field.Value;
                                    int length = sourceproj.TeamProjectName.Length;
                                    string itPathNew = destinationproj.TeamProjectName + itPath.Substring(length);
                                    newWorkItem.Fields[field.Name].Value = itPathNew;
                                }
                                catch (Exception)
                                {
                                }
                            }
                        }
                        //Add values to mapped fields
                        else if (fieldMap.ContainsKey(field.Name))
                        {
                            try
                            {
                                newWorkItem.Fields[(string)fieldMap[field.Name]].Value = field.Value;
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }

                    }
                    ArrayList array = newWorkItem.Validate();
                    foreach (Field item in array)
                    {
                        logger.Info(String.Format("Work item {0} Validation Error in field: {1}  : {2}", item.Name, newWorkItem.Fields[item.Name].Value));
                    }
                    if (array.Count == 0)
                    {
                        SaveAttachments(testcase.TestCase.WorkItem);
                        UploadAttachments(newWorkItem, testcase.TestCase.WorkItem);
                        newWorkItem.Save();
                        ITestCase tc = destinationproj.TestCases.Find(newWorkItem.Id);
                        destinationsuite.Entries.Add(tc);
                        TestActionCollection testActionCollection = testcase.TestCase.Actions;
                        foreach (var item in testActionCollection)
                        {
                            item.CopyToNewOwner(tc);                            
                        }
                        tc.Save();
                    }
                   
                }
                catch (Exception ex)
                {
                    logger.Info("Error retrieving Test case  " + testcase.TestCase.WorkItem.Id + ": " + testcase.Title + ex.Message);
                }
            }
        }

        private void SaveAttachments(WorkItem wi)
        {
            System.Net.WebClient webClient = new System.Net.WebClient();
            foreach (Attachment att in wi.Attachments)
            {
                try
                {
                    String path = @"Attachments\" + wi.Id;
                    bool folderExists = Directory.Exists(path);
                    if (!folderExists)
                    {
                        Directory.CreateDirectory(path);
                    }
                    if (!File.Exists(path + "\\" + att.Name))
                    {
                        webClient.DownloadFile(att.Uri, path + "\\" + att.Name);
                    }
                    else
                    {
                        webClient.DownloadFile(att.Uri, path + "\\" + att.Id + "_" + att.Name);
                    }

                }
                catch (Exception)
                {
                    logger.Info("Error downloading attachment for work item : " + wi.Id + " Type: " + wi.Type.Name);
                }

            }
        }

        private void UploadAttachments(WorkItem workItem, WorkItem workItemOld)
        {
            AttachmentCollection attachmentCollection = workItemOld.Attachments;
            foreach (Attachment att in attachmentCollection)
            {
                string comment = att.Comment;
                string name = @"Attachments\" + workItemOld.Id + "\\" + att.Name;
                string nameWithID = @"Attachments\" + workItemOld.Id + "\\" + att.Id + "_" + att.Name;
                try
                {
                    if (File.Exists(nameWithID))
                    {
                        workItem.Attachments.Add(new Attachment(nameWithID, comment));
                    }
                    else
                    {
                        workItem.Attachments.Add(new Attachment(name, comment));
                    }
                }
                catch (Exception ex)
                {
                    logger.ErrorFormat("Error saving attachment: {0} for workitem: {1}", att.Name, workItemOld.Id);
                    logger.Error("Error detail: ", ex);
                }
            }
        }

        private Hashtable ListToTable(List<object> map)
        {
            Hashtable table = new Hashtable();
            if (map != null)
            {
                foreach (object[] item in map)
                {
                    try
                    {
                        table.Add((string)item[0], (string)item[1]);
                    }
                    catch (Exception ex)
                    {
                        logger.Error("Error in ListToTable", ex);
                    }
                }
            }
            return table;
        }
    }


    
}
