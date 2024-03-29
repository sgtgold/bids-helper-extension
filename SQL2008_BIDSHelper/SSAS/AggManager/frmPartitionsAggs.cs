
/*============================================================================
  File:    frmPartitionAggs.cs

  Summary: Contains the form to display information about aggregation sizes

  Summary: Contains the form to add aggregations based on informaion in the Query Log

           Part of Aggregation Manager 

  Date:    January 2007
------------------------------------------------------------------------------
  This file is part of the Microsoft SQL Server Code Samples.

  Copyright (C) Microsoft Corporation.  All rights reserved.

  This source code is intended only as a supplement to Microsoft
  Development Tools and/or on-line documentation.  See these other
  materials for detailed information regarding Microsoft code samples.

  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
  PARTICULAR PURPOSE.
============================================================================*/
/*
 * This file has been incorporated into BIDSHelper. 
 *    http://www.codeplex.com/BIDSHelper
 * and may have been altered from the orginal version which was released 
 * as a Microsoft sample.
 * 
 * The official version can be found on the sample website here: 
 * http://www.codeplex.com/MSFTASProdSamples                                   
 *                                                                             
 ============================================================================*/
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.AnalysisServices;
using Microsoft.AnalysisServices.AdomdClient;

namespace AggManager
{
    public partial class PartitionsAggsForm : Form
    {
        private MeasureGroup mg1;
        private Partition part1;
        private AggregationDesign aggDes;
        DataSet partitionDetails;

        public void Init(MeasureGroup mg, 
            string strParition,
            EnvDTE.ProjectItem projItem)
        {
            try
            {
                bool IsOnlineMode = false;
                Cube selectedCube = projItem.Object as Cube;
                string serverName = "";
                string databaseName = "";
                if ((selectedCube != null) && (selectedCube.ParentServer != null))
                {
                    // if we are in Online mode there will be a parent server
                    serverName = selectedCube.ParentServer.Name;
                    databaseName = selectedCube.Parent.Name;
                    IsOnlineMode = true;
                }
                else
                {
                    // if we are in Project mode we will use the server name from 
                    // the deployment settings
                    DeploymentSettings deploySet = new DeploymentSettings(projItem);
                    serverName = deploySet.TargetServer;
                    databaseName = deploySet.TargetDatabase; //use the target database instead of selectedCube.Parent.Name because selectedCube.Parent.Name only reflects the last place it was deployed to, and we want the user to be able to use the deployment settings to control which deployed server/database to check against
                }
                mg1 = mg;
                part1 = mg.Partitions.FindByName(strParition);
                aggDes = part1.AggregationDesign;
                this.Text = " Aggregation sizes for partition " + strParition;

                //lblSize.Text = part1.EstimatedRows.ToString() + " records"; //base this not on estimated rows but on actual rows... see below
                lablPartName.Text = strParition;

                txtServerNote.Text = string.Format("Note: The Partition size details have been taken from the currently deployed '{1}' database on the '{0}' server, " +
                "which is the one currently configured as the deployment target.", serverName, databaseName);
                txtServerNote.Visible = !IsOnlineMode;

                //--------------------------------------------------------------------------------
                // Open ADOMD connection to the server and issue DISCOVER_PARTITION_STAT request to get aggregation sizes
                //--------------------------------------------------------------------------------
                AdomdConnection adomdConnection = new AdomdConnection("Data Source=" + serverName);
                adomdConnection.Open();
                partitionDetails = adomdConnection.GetSchemaDataSet(AdomdSchemaGuid.PartitionStat, new object[] { databaseName, mg1.Parent.Name, mg1.Name, strParition });

                DataColumn colItem1 = new DataColumn("Percentage", Type.GetType("System.String"));
                partitionDetails.Tables[0].Columns.Add(colItem1);

                AddGridStyle();

                dataGrid1.DataSource = partitionDetails.Tables[0];

                long iPartitionRowCount = 0;
                if (partitionDetails.Tables[0].Rows.Count > 0)
                {
                    iPartitionRowCount = Convert.ToInt64(partitionDetails.Tables[0].Rows[0]["AGGREGATION_SIZE"]);
                }
                lblSize.Text = iPartitionRowCount + " records";

                double ratio = 0;
                foreach (DataRow row in partitionDetails.Tables[0].Rows)
                {
                    ratio = 100.0 * ((long)row["AGGREGATION_SIZE"] / (double)iPartitionRowCount);
                    row["Percentage"] = ratio.ToString("#0.00") + "%";
                }

                CurrencyManager cm = (CurrencyManager)this.BindingContext[dataGrid1.DataSource, dataGrid1.DataMember];
                ((DataView)cm.List).AllowNew = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
                try
                {
                    this.Close();
                }
                catch { }
            }
        }

        public PartitionsAggsForm()
        {
            InitializeComponent();
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Would you like to save the aggregation design: " + aggDes.Name + "?\r\nNote, aggregations not found in the list will be deleted from the aggregation design.\r\n\r\nNOTE, THIS LIST REFLECTS WHAT HAS BEEN DEPLOYED TO THE SERVER AND PROCESSED, SO IT MAY NOT HAVE NEWLY DESIGNED AGGREGATIONS!", "Save Message", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                    return;

                bool boolAggFound = false;

                for (int i = 0; aggDes.Aggregations.Count > i; i++)
                {
                    boolAggFound = false;

                    foreach (DataRow dRow in partitionDetails.Tables[0].Rows)
                        if (dRow["AGGREGATION_NAME"].ToString() == aggDes.Aggregations[i].Name) boolAggFound = true;

                    if (!boolAggFound)
                    {
                        aggDes.Aggregations.Remove(aggDes.Aggregations[i].Name);
                        i--;
                    }
                }

                //MessageBox.Show("Aggregation design: " + aggDes.Name + "  has been updated with " + aggDes.Aggregations.Count.ToString() + " aggregations ");

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void AddGridStyle()
        {

            int iWidth0 = 300;
            Graphics Graphics = dataGrid1.CreateGraphics();

            if (partitionDetails.Tables[0].Rows.Count > 0)
            {
                int iColWidth = (int)(Graphics.MeasureString
                    (partitionDetails.Tables[0].Rows[0].ItemArray[0].ToString(),
                    dataGrid1.Font).Width);
                iWidth0 = (int)System.Math.Max(iWidth0, iColWidth);
            }

            DataGridTableStyle myGridStyle = new DataGridTableStyle();
            myGridStyle.MappingName = "rowsettable";

            DataGridTextBoxColumn nameColumnStyle = new DataGridTextBoxColumn();
            nameColumnStyle.MappingName = "AGGREGATION_NAME";
            nameColumnStyle.HeaderText = "Aggregation Name";
            nameColumnStyle.Width = iWidth0 + 10;
            myGridStyle.GridColumnStyles.Add(nameColumnStyle);

            DataGridTextBoxColumn nameColumnStyle1 = new DataGridTextBoxColumn();
            nameColumnStyle1.MappingName = "AGGREGATION_SIZE";
            nameColumnStyle1.HeaderText = "Records";
            nameColumnStyle1.Width = 70;
            myGridStyle.GridColumnStyles.Add(nameColumnStyle1);

            DataGridTextBoxColumn nameColumnStyle2 = new DataGridTextBoxColumn();
            nameColumnStyle2.MappingName = "Percentage";
            nameColumnStyle2.HeaderText = "Percentage";
            nameColumnStyle2.Width = 100;
            myGridStyle.GridColumnStyles.Add(nameColumnStyle2);

            dataGrid1.TableStyles.Add(myGridStyle);

        }


    }
}