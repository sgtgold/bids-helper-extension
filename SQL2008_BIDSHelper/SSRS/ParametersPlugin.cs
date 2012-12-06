using System;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using System.Xml;
using System.Text;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Data;

namespace BIDSHelper.SSRS
{
    public class ParametersPlugin : BIDSHelperPluginBase
    {
        public ParametersPlugin(Connect con, DTE2 appObject, AddIn addinInstance)
            : base(con, appObject, addinInstance)
        {
        }

        public override string ShortName
        {
            get { return "Parameters"; }
        }

        public override int Bitmap
        {
            get { return 543; }
        }

        public override string ButtonText
        {
            get { return "Parameters Details ..."; }
        }

        public override string ToolTip
        {
            get { return string.Empty; } //not used anywhere
        }

        public override string MenuName
        {
            get { return "Project,Solution,File"; }
        }

        public override bool ShouldPositionAtEnd
        {
            get { return true; }
        }

        /// <summary>
        /// Gets the name of the friendly name of the plug-in.
        /// </summary>
        /// <value>The friendly name.</value>
        public override string FeatureName
        {
            get { return "Unused Report Datasets"; }
        }

        /// <summary>
        /// Gets the Url of the online help page for this plug-in.
        /// </summary>
        /// <value>The help page Url.</value>
        public override string  HelpUrl
        {
	        get { return this.GetCodePlexHelpUrl("Dataset Usage Reports"); }
        }

        /// <summary>
        /// Gets the feature category used to organise the plug-in in the enabled features list.
        /// </summary>
        /// <value>The feature category.</value>
        public override BIDSFeatureCategories FeatureCategory
        {
            get { return BIDSFeatureCategories.SSRS; }
        }

        /// <summary>
        /// Gets the full description used for the features options dialog.
        /// </summary>
        /// <value>The description.</value>
        public override string FeatureDescription
        {
            get { return "To-do."; }
        }

        /// <summary>
        /// Determines if the command should be displayed or not.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public override bool DisplayCommand(UIHierarchyItem item)
        {
            try
            {
                UIHierarchy solExplorer = this.ApplicationObject.ToolWindows.SolutionExplorer;
                if (((System.Array)solExplorer.SelectedItems).Length != 1)
                    return false;

                UIHierarchyItem hierItem = ((UIHierarchyItem)((System.Array)solExplorer.SelectedItems).GetValue(0));
                SolutionClass solution = hierItem.Object as SolutionClass;
                Project project = hierItem.Object as Project;
                if (project == null && hierItem.Object is ProjectItem)
                {
                    ProjectItem pi = hierItem.Object as ProjectItem;
                    project = pi.SubProject;
                                       
                }
                if (project != null)
                {
                  if (GetRdlFilesInProjectItems(project.ProjectItems, true).Length > 0)
                        return true;
                }
                else if (solution != null)
                {
                    foreach (Project p in solution.Projects)
                    {
                        if (GetRdlFilesInProjectItems(p.ProjectItems, true).Length > 0)
                            return true;
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        public static string[] GetRdlFilesInProjectItems(ProjectItems pis, bool bGetRDLC)
        {
            if (pis == null) return new string[] {};

            List<string> lst = new List<string>();
            foreach (ProjectItem pi in pis)
            {
                if (pi.SubProject != null)
                {
                    lst.AddRange(GetRdlFilesInProjectItems(pi.SubProject.ProjectItems, bGetRDLC));
                }
                else if (pi.Name.ToLower().EndsWith(".rdl") || (bGetRDLC && pi.Name.ToLower().EndsWith(".rdlc")))
                {
                    lst.Add(pi.get_FileNames(1));
                }
                lst.AddRange(GetRdlFilesInProjectItems(pi.ProjectItems, bGetRDLC));
            }
            return lst.ToArray();
        }

        public override void Exec()
        {
            try
            {
                TestWpf.MainWindow win = new TestWpf.MainWindow();
                win.ShowDialog();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }


            //ScanReports(true);
        }

        //protected void ScanReports(bool LookForUnusedDatasets)
        //{
        //    string sCurrentFile = string.Empty;
        //    try
        //    {
        //        UIHierarchy solExplorer = this.ApplicationObject.ToolWindows.SolutionExplorer;
        //        UIHierarchyItem hierItem = ((UIHierarchyItem)((System.Array)solExplorer.SelectedItems).GetValue(0));
        //        SolutionClass solution = hierItem.Object as SolutionClass;

        //        List<string> lstRdls = new List<string>();
        //        if (hierItem.Object is Project)
        //        {
        //            Project p = (Project)hierItem.Object;
        //            lstRdls.AddRange(GetRdlFilesInProjectItems(p.ProjectItems, true));
        //        }
        //        else if (hierItem.Object is ProjectItem)
        //        {
        //            ProjectItem pi = hierItem.Object as ProjectItem;
        //            Project p = pi.SubProject;
        //            lstRdls.AddRange(GetRdlFilesInProjectItems(p.ProjectItems, true));
        //        }
        //        else if (solution != null)
        //        {
        //            foreach (Project p in solution.Projects)
        //            {
        //                lstRdls.AddRange(GetRdlFilesInProjectItems(p.ProjectItems, true));
        //            }
        //        }

        //        List<UsedRsDataSets.RsDataSetUsage> lstDataSets = new List<UsedRsDataSets.RsDataSetUsage>();
        //        foreach (string file in lstRdls)
        //        {
        //            sCurrentFile = file;
        //            UsedRsDataSets urds = new UsedRsDataSets(file);
        //            foreach (UsedRsDataSets.RsDataSet ds in urds.DataSets)
        //            {
        //                if (LookForUnusedDatasets && ds.Usages.Count == 0)
        //                {
        //                    UsedRsDataSets.RsDataSetUsage u = new UsedRsDataSets.RsDataSetUsage();
        //                    u.ReportName = ds.ReportName;
        //                    u.DataSetName = ds.DataSetName;
        //                    lstDataSets.Add(u);
        //                }
        //                else if (!LookForUnusedDatasets && ds.Usages.Count > 0)
        //                {
        //                    foreach (string usage in ds.Usages)
        //                    {
        //                        UsedRsDataSets.RsDataSetUsage u = new UsedRsDataSets.RsDataSetUsage();
        //                        u.ReportName = ds.ReportName;
        //                        u.DataSetName = ds.DataSetName;
        //                        u.Usage = usage;
        //                        lstDataSets.Add(u);
        //                    }
        //                }
        //            }
        //        }

        //        if (lstDataSets.Count == 0)
        //        {
        //            if (LookForUnusedDatasets)
        //                MessageBox.Show("All datasets are in use.", "BIDS Helper Unused Datasets Report");
        //            else
        //                MessageBox.Show("No datasets found.", "BIDS Helper Used Datasets Report");
        //        }
        //        else
        //        {
        //            ReportViewerForm frm = new ReportViewerForm();
        //            frm.ReportBindingSource.DataSource = lstDataSets;
        //            if (LookForUnusedDatasets)
        //            {
        //                frm.Report = "SSRS.UnusedDatasets.rdlc";
        //                frm.Caption = "Unused Datasets Report";
        //            }
        //            else
        //            {
        //                frm.Report = "SSRS.UsedDatasets.rdlc";
        //                frm.Caption = "Used Datasets Report";
        //            }
        //            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource1 = new Microsoft.Reporting.WinForms.ReportDataSource();
        //            reportDataSource1.Name = "BIDSHelper_SSRS_RsDataSetUsage";
        //            reportDataSource1.Value = frm.ReportBindingSource;
        //            frm.ReportViewerControl.LocalReport.DataSources.Add(reportDataSource1);

        //            frm.WindowState = FormWindowState.Maximized;
        //            frm.Show();
        //        }
        //    }
        //    catch (System.Exception ex)
        //    {
        //        string sError = string.Empty;
        //        if (!string.IsNullOrEmpty(sCurrentFile)) sError += "Error while scanning report: " + sCurrentFile + "\r\n";
        //        while (ex != null)
        //        {
        //            sError += ex.Message + "\r\n" + ex.StackTrace + "\r\n\r\n";
        //            ex = ex.InnerException;
        //        }
        //        MessageBox.Show(sError);
        //    }
        //}

    }

 
 
}
