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
    public class ReportParametersPlugin : BIDSHelperPluginBase
    {
        public ReportParametersPlugin(Connect con, DTE2 appObject, AddIn addinInstance)
            : base(con, appObject, addinInstance)
        {
        }

        public override string ShortName
        {
            get { return "ReportParameters"; }
        }

        public override int Bitmap
        {
            get { return 543; }
        }

        public override string ButtonText
        {
            get { return "Report Parameters Details..."; }
        }

        public override string ToolTip
        {
            get { return "my Tooltip"; } //not used anywhere
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
            get { return this.GetCodePlexHelpUrl("Report Parameters Details"); }
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
            get { return "Provides a overview of report parameters."; }
        }
        
        private string[] SSRS_FILE_EXTENSIONS = { ".rdl" };

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
                string sFileName = ((ProjectItem)hierItem.Object).Name.ToLower();
                foreach (string extension in SSRS_FILE_EXTENSIONS)
                {
                    if (sFileName.EndsWith(extension))
                        return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        public override void Exec()
        {
            try
            {
            UIHierarchy solExplorer = this.ApplicationObject.ToolWindows.SolutionExplorer;
            UIHierarchyItem hierItem = (UIHierarchyItem)((System.Array)solExplorer.SelectedItems).GetValue(0);
            ProjectItem projItem = (ProjectItem)hierItem.Object;

            string sFileName = ((ProjectItem)hierItem.Object).Name.ToLower();
            
            if (this.ApplicationObject.ActiveDocument != null)
            {
                //this.ApplicationObject.ActiveDocument;
                //this.ApplicationObject.ActiveDocument.FullName        cesta k souboru
                //this.ApplicationObject.ActiveDocument.Saved           false když není uložený
            }

            ReportParametersUI.MainWindow w = new ReportParametersUI.MainWindow();

            MessageBox.Show("instance hotova");

            w.ShowDialog();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

    }

    

}
