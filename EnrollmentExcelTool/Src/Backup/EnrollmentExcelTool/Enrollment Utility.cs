using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;


namespace EnrollmentExcelTool
{
    public partial class Form1 : Form
    {
        private MasterFileInfo masterFileInfo;
        private SourceFileInfo sourceFileInfo;
        private ProfileFileInfo profileFileInfo;
       
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (cmbSelectOption.SelectedIndex == -1)
            {
                cmbSelectOption.SelectedIndex = 0;
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            btnRun.Enabled = false;

            string srcFileName = string.Empty;
            string targetFileName = string.Empty;

            srcFileName = txtMasterFile.Text;
            targetFileName = txtSourceFile.Text;

            if (string.IsNullOrEmpty(srcFileName) || string.IsNullOrEmpty(targetFileName))
            {
                MessageBox.Show("Please Select Files for encoding", "Select Files");
                btnRun.Enabled = true;
                return;
            }

            
            masterFileInfo = new MasterFileInfo();
            sourceFileInfo = new SourceFileInfo();
            profileFileInfo = new ProfileFileInfo();
            string errorList = string.Empty;

            if (cmbSelectOption.SelectedItem.ToString() == "Master To Source")
            {
                //In this case srcFileName is Master File  and targetFileName is Source File
                DialogResult response = MessageBox.Show("Are you sure, the following selected files are correct and of same model?" + Environment.NewLine + Environment.NewLine + "Master file: " + srcFileName + Environment.NewLine + "and " + Environment.NewLine + " Source file: " + targetFileName, "Confirm file selection", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (response != DialogResult.Yes)
                {
                    btnRun.Enabled = true;
                    return;
                }
                
                errorList = validateMasterFile(srcFileName);

                errorList = errorList +  validateSourceFile(targetFileName);

                if (masterFileInfo.isValid==false || sourceFileInfo.isValid == false)
                {
                    MessageBox.Show(errorList, "Errors",MessageBoxButtons.OK, MessageBoxIcon.Error);
                    btnRun.Enabled = true;
                    return;
                }

                //Encode Source Files CAMP CODE based on Description, Position and Part Number
                EncodeSourceFileCampCodesBasedOnDPP(srcFileName, targetFileName);

                //Encode Source Files CAMP CODE based on Description and Position 
                EncodeSourceFileCampCodesBasedOnDP(srcFileName, targetFileName);

                //Encode Source Files CAMP CODE based on Position and Part Number
                EncodeSourceFileCampCodesBasedOnPP(srcFileName, targetFileName);

                //Encode Source Files CAMP CODE based on Part Number
                EncodeSourceFileCampCodesBasedOnPartNumber(srcFileName, targetFileName);

                //Encode Source Files CAMP CODE based on Description
                EncodeSourceFileCampCodesBasedOnDescription(srcFileName, targetFileName);

                //Update Souce file's Duplicate column to 'YES' for duplicate CAMP CODES 
                UpdateSourceDataForDuplicates(targetFileName);
                MessageBox.Show("Source File Encoding Completed", "Completed");
            }
            else
            {
                //In this case srcFileName is Source File  and targetFileName is Profile File
                DialogResult response = MessageBox.Show("Are you sure, the following selected files are correct and of same model?" + Environment.NewLine + Environment.NewLine + "Source file: " + srcFileName + Environment.NewLine + "and " + Environment.NewLine + " Profile file: " + targetFileName, "Confirm file selection", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (response != DialogResult.Yes)
                {
                    btnRun.Enabled = true;
                    return;
                }
                
                errorList = validateSourceFileForProfileEncoding(srcFileName);

                errorList = errorList + validateProfileFile(targetFileName);

                if (sourceFileInfo.isValid == false || profileFileInfo.isValid == false)
                {
                    MessageBox.Show(errorList, "Errors", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    btnRun.Enabled = true;
                    return;
                }
                //Encode PN ON and SN ON of Profile file based on CAMP CODE
                EncodeProfileData(srcFileName, targetFileName);

                //Update Souce file's Duplicate column to 'YES' for duplicate CAMP CODES 
                UpdateSourceDataForDuplicates(srcFileName);
                MessageBox.Show("Profile Spreadsheet Enrolling Completed");
            }
            btnRun.Enabled = true;
        }

        private void cmbSelectOption_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbSelectOption.SelectedItem.ToString() == "Master To Source")
            {
                grpMaster.Text = "Master To Source";
                btnSelectMaster.Text = "Select Master File";
                btnSource.Text = "Select Source File";
            }
            else
            {
                grpMaster.Text = "Source To Profile";
                btnSelectMaster.Text = "Select Source File";
                btnSource.Text = "Select Profile File";
            }
        }

        private void btnSource_Click(object sender, EventArgs e)
        {
            OpenFileDialog fDialog = new OpenFileDialog();
            fDialog.Title = "Open Source File";
            fDialog.Filter = "Excel Files|*.xls";
            
            if (fDialog.ShowDialog() == DialogResult.OK)
            {
                txtSourceFile.Text = fDialog.FileName.ToString();
            }
        }

        private void btnSelectMaster_Click(object sender, EventArgs e)
        {
            OpenFileDialog fDialog = new OpenFileDialog();
            fDialog.Title = "Open Master Reference File";
            fDialog.Filter = "Excel Files|*.xls";
            
            if (fDialog.ShowDialog() == DialogResult.OK)
            {
                txtMasterFile.Text = fDialog.FileName.ToString();
            }
        }

        public void EncodeProfileData(string sourceFileName, string profileFileName)
        {
            System.Data.DataTable dt = null;
            
            OleDbConnection SrcConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @sourceFileName + ";Extended Properties=" + (char)34 + "Excel 8.0;HDR=Yes" + (char)34);
            SrcConn.Open();

            string sql = string.Empty;
            string searchString = string.Empty;
            string CAMPCODE = string.Empty;
            int result = 0;
            dt = SrcConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string sheetName;
            sheetName = dt.Rows[0]["TABLE_NAME"].ToString();

            sql = "select * from [" + sheetName + "] WHERE [CAMP CODE] is NOT Null AND RTRIM([CAMP CODE]) <> ''";            
            OleDbCommand srcCommand = new OleDbCommand();
            OleDbCommand srcUpdCommand = new OleDbCommand();
            srcUpdCommand.Connection = SrcConn;
            srcCommand.CommandText = sql;
            srcCommand.Connection = SrcConn;
            OleDbDataReader dr = srcCommand.ExecuteReader();
          
            while (dr.Read())
            {
                if (dr[0].ToString() == "" && dr[1].ToString() == "" && dr[3].ToString() == "" &&
                    dr[4].ToString() == "" && dr[5].ToString() == "" && dr[6].ToString() == "")
                {
                    break;
                }
                CAMPCODE = dr["CAMP CODE"].ToString().Trim();
                searchString = CAMPCODE + "PN" + "SN-UNKNOWN";               
                searchString = searchString.Replace("'", "''");
                
                if (CAMPCODE != "")
                {                    
                    try
                    {
                        OleDbConnection profileConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @profileFileName + ";Extended Properties=" + (char)34 + "Excel 8.0;HDR=YES" + (char)34);

                        OleDbCommand profileCommand = new OleDbCommand();
                        
                        profileConn.Open();
                        dt = profileConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        string sheetName1;
                        sheetName1 = dt.Rows[0]["TABLE_NAME"].ToString();
                        
                        profileCommand.Connection = profileConn;

                        string ssql = string.Empty;
                        ssql = "Select Top 1 * FROM [" + sheetName1 + "]  WHERE RTRIM([CAMP CODE]) + RTRIM([PN ON]) + RTRIM([SN ON]) = '" + searchString + "' ORDER BY [Id]";

                        profileCommand.CommandText = ssql;
                        OleDbDataReader selectDr = profileCommand.ExecuteReader();
                        string profileUniqueId = string.Empty;
                        while (selectDr.Read())
                        {
                            profileUniqueId = selectDr["Id"].ToString();
                        }
                        selectDr.Close();

                        if (profileUniqueId != "")
                        {
                            sql = "Update [" + sheetName1 + "] SET [PN ON] ='" + dr["Part Number"].ToString().Trim() + "', [SN ON] ='" + dr["Serial Number"].ToString().Trim() + "' , [WORK PERFORMED ICAO] ='" + dr["COMP_CITY_INSTLD"].ToString().Trim() + "' ,[PAGE] ='" + dr["PAGE"].ToString().Trim() + "' , [Status] = 'EXACT' WHERE RTRIM([Id])  = '" + profileUniqueId + "' ";
                        }
                        else
                        {
                            sql = "Update [" + sheetName1 + "] SET [PN ON] ='" + dr["Part Number"].ToString().Trim() + "', [SN ON] ='" + dr["Serial Number"].ToString().Trim() + "' , [WORK PERFORMED ICAO] ='" + dr["COMP_CITY_INSTLD"].ToString().Trim() + "' ,[PAGE] ='" + dr["PAGE"].ToString().Trim() + "', [Status] = 'EXACT1' WHERE RTRIM([CAMP CODE]) + RTRIM([PN ON]) + RTRIM([SN ON]) = '" + searchString + "' ";
                        }
                        profileCommand.CommandText = sql;
                        result = profileCommand.ExecuteNonQuery();
                 
                        
                        profileConn.Close();
                        sql = "";
                        if (result >0)
                        {
                            sql = "Update [" + sheetName + "] SET [ADDRESSED] ='YES' WHERE [Id]='" + dr["Id"] + "'";                             
                            srcUpdCommand.CommandText = sql;
                            result = srcUpdCommand.ExecuteNonQuery();
                        }
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show (ex.Message);
                    }
                }   
            
            }
            SrcConn.Close();            
        }

        public void EncodeSourceFileCampCodesBasedOnDPP(string masterFileName, string sourceFileName)
        {
            System.Data.DataTable dt = null;            
            OleDbConnection masterConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @masterFileName + ";Extended Properties=" + (char)34 + "Excel 8.0;HDR=Yes" + (char)34);
            masterConn.Open();

            string sql = string.Empty;
            string searchString = string.Empty;
            string masterCampCode = string.Empty;
            string enCodingCondition = string.Empty;
            int result = 0;
            dt = masterConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string sheetName;
            sheetName = dt.Rows[0]["TABLE_NAME"].ToString();

            sql = "select * from [" + sheetName + "] WHERE [Encoded]<>'YES' OR [Encoded] is NULL ORDER BY [Id]";
            OleDbCommand masterCommand = new OleDbCommand();
            OleDbCommand masterUpdCommand = new OleDbCommand();
            masterUpdCommand.Connection = masterConn;
            masterCommand.CommandText = sql;
            masterCommand.Connection = masterConn;
            OleDbDataReader dr = masterCommand.ExecuteReader();
            
            while (dr.Read())
            {
                if (dr[0].ToString() == "" && dr[1].ToString() == "" && dr[3].ToString() == "" &&
                    dr[4].ToString() == "" && dr[5].ToString() == "" && dr[6].ToString() == "")
                {
                    break;
                }
                searchString = string.Empty;
                if (!string.IsNullOrEmpty(dr["CAMP CODE"].ToString()))
                {
                    masterCampCode = dr["CAMP CODE"].ToString().Trim();
                }
                else
                {
                    masterCampCode = "";
                }
                //Description
                if (!string.IsNullOrEmpty(dr["Description"].ToString()))
                {
                    searchString = dr["Description"].ToString().Trim();
                }
                else
                {
                    searchString = "";
                }
                //Position
                if (!string.IsNullOrEmpty(dr["Position"].ToString()))
                {
                    searchString = searchString + dr["Position"].ToString().Trim();
                }
                else
                {
                    searchString = searchString + "";
                }
                //Part Number
                if (!string.IsNullOrEmpty(dr["Part Number"].ToString()))
                {
                    searchString = searchString + dr["Part Number"].ToString().Trim();
                }
                else
                {
                    searchString = searchString + "";
                }                
                
                searchString = searchString.Replace("'", "''");
                enCodingCondition = "DPP";

                if (masterCampCode != "")
                {
                    
                    try
                    {
                        OleDbConnection srcConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @sourceFileName + ";Extended Properties=" + (char)34 + "Excel 8.0;HDR=YES" + (char)34);

                        OleDbCommand srcCommand = new OleDbCommand();

                        srcConn.Open();
                        dt = srcConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        string sheetName1;
                        sheetName1 = dt.Rows[0]["TABLE_NAME"].ToString();

                        srcCommand.Connection = srcConn;

                        string ssql = string.Empty;
                        ssql = "Select Top 1 * FROM [" + sheetName1 + "]  WHERE RTRIM([Description]) + RTRIM([Position]) + LTRIM([Part Number]) = '" + searchString + "' AND ( [CAMP CODE] is Null OR RTRIM([CAMP CODE]) ='') ORDER BY [Id]";

                        srcCommand.CommandText = ssql;
                        OleDbDataReader selectDr = srcCommand.ExecuteReader();
                        string srcUniqueId = string.Empty;
                        while (selectDr.Read())
                        {
                            srcUniqueId = selectDr["Id"].ToString();
                        }
                        selectDr.Close();
                        if (srcUniqueId != "")
                        {
                            //sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "', [Encoding Condition] ='" + enCodingCondition + "' WHERE RTRIM([Id]) ='" + srcUniqueId + "' ";
                            sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "' WHERE RTRIM([Id]) ='" + srcUniqueId + "' ";
                        }
                        else
                        {
                            //sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "', [Encoding Condition] ='" + enCodingCondition + "' WHERE RTRIM([Description]) + RTRIM([Position]) + LTRIM([Part Number]) = '" + searchString + "' AND ( [CAMP CODE] is Null OR RTRIM([CAMP CODE]) ='') ";
                            sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "' WHERE RTRIM([Description]) + RTRIM([Position]) + LTRIM([Part Number]) = '" + searchString + "' AND ( [CAMP CODE] is Null OR RTRIM([CAMP CODE]) ='') ";
                        }
                        srcCommand.CommandText = sql;
                        result = srcCommand.ExecuteNonQuery();                        

                        srcConn.Close();
                        sql = "";
                        if (result > 0)
                        {
                            sql = "Update [" + sheetName + "] SET [Encoded] ='YES' WHERE  [Id]= '" + dr["Id"].ToString().Trim() + "'";
                            masterUpdCommand.CommandText = sql;
                            result = masterUpdCommand.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }

            }
            masterConn.Close();
           
        }

        public void EncodeSourceFileCampCodesBasedOnDP(string masterFileName, string sourceFileName)
        {
            System.Data.DataTable dt = null;
            OleDbConnection masterConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @masterFileName + ";Extended Properties=" + (char)34 + "Excel 8.0;HDR=Yes" + (char)34);
            masterConn.Open();

            string sql = string.Empty;
            string searchString = string.Empty;
            string masterCampCode = string.Empty;
            string enCodingCondition = string.Empty;
            int result = 0;
            dt = masterConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string sheetName;
            sheetName = dt.Rows[0]["TABLE_NAME"].ToString();

            sql = "select * from [" + sheetName + "] WHERE [Encoded]<>'YES' OR [Encoded] is NULL ORDER BY [Id]";
            OleDbCommand masterCommand = new OleDbCommand();
            OleDbCommand masterUpdCommand = new OleDbCommand();
            masterUpdCommand.Connection = masterConn;
            masterCommand.CommandText = sql;
            masterCommand.Connection = masterConn;
            OleDbDataReader dr = masterCommand.ExecuteReader();
            
            while (dr.Read())
            {
                if (dr[0].ToString() == "" && dr[1].ToString() == "" && dr[3].ToString() == "" &&
                    dr[4].ToString() == "" && dr[5].ToString() == "" && dr[6].ToString() == "")
                {
                    break;
                }
                searchString = string.Empty;
                if (!string.IsNullOrEmpty(dr["CAMP CODE"].ToString()))
                {
                    masterCampCode = dr["CAMP CODE"].ToString().Trim();
                }
                else
                {
                    masterCampCode = "";
                }
                //Description
                if (!string.IsNullOrEmpty(dr["Description"].ToString()))
                {
                    searchString = dr["Description"].ToString().Trim();
                }
                else
                {
                    searchString = "";
                }
                //Position
                if (!string.IsNullOrEmpty(dr["Position"].ToString()))
                {
                    searchString = searchString + dr["Position"].ToString().Trim();
                }
                else
                {
                    searchString = searchString + "";
                }                

                searchString = searchString.Replace("'", "''");
                
                enCodingCondition = "DP";

                if (masterCampCode != "")
                {                   
                    try
                    {
                        OleDbConnection srcConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @sourceFileName + ";Extended Properties=" + (char)34 + "Excel 8.0;HDR=YES" + (char)34);

                        OleDbCommand srcCommand = new OleDbCommand();

                        srcConn.Open();
                        dt = srcConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        string sheetName1;
                        sheetName1 = dt.Rows[0]["TABLE_NAME"].ToString();

                        srcCommand.Connection = srcConn;

                        string ssql = string.Empty;
                        ssql = "Select Top 1 * FROM [" + sheetName1 + "]  WHERE RTRIM([Description]) + RTRIM([Position])  = '" + searchString + "' AND ( [CAMP CODE] is Null OR RTRIM([CAMP CODE]) ='') ORDER BY [Id]";

                        srcCommand.CommandText = ssql;
                        OleDbDataReader selectDr = srcCommand.ExecuteReader();
                        string srcUniqueId = string.Empty;
                        while (selectDr.Read())
                        {
                            srcUniqueId = selectDr["Id"].ToString().Trim();
                        }
                        selectDr.Close();

                        if (srcUniqueId != "")
                        {
                            //sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "', [Encoding Condition] ='" + enCodingCondition + "' WHERE  RTRIM([Id]) ='" + srcUniqueId + "' ";
                            sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "' WHERE  RTRIM([Id]) ='" + srcUniqueId + "' ";
                        }
                        else
                        {
                            //sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "', [Encoding Condition] ='" + enCodingCondition + "' WHERE RTRIM([Description]) + RTRIM([Position]) = '" + searchString + "' AND ( [CAMP CODE] is Null OR RTRIM([CAMP CODE]) ='') ";
                            sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "' WHERE RTRIM([Description]) + RTRIM([Position]) = '" + searchString + "' AND ( [CAMP CODE] is Null OR RTRIM([CAMP CODE]) ='') ";
                        }
                        srcCommand.CommandText = sql;
                        result = srcCommand.ExecuteNonQuery();

                        srcConn.Close();
                        sql = "";
                        if (result > 0)
                        {
                            sql = "Update [" + sheetName + "] SET [Encoded] ='YES' WHERE  [Id]= '" + dr["Id"].ToString().Trim() + "'";
                            masterUpdCommand.CommandText = sql;
                            result = masterUpdCommand.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }

            }
            masterConn.Close();
        }

        public void EncodeSourceFileCampCodesBasedOnPP(string masterFileName, string sourceFileName)
        {
            System.Data.DataTable dt = null;
            OleDbConnection masterConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @masterFileName + ";Extended Properties=" + (char)34 + "Excel 8.0;HDR=Yes" + (char)34);
            masterConn.Open();

            string sql = string.Empty;
            string searchString = string.Empty;
            string masterCampCode = string.Empty;
            string enCodingCondition = string.Empty;
            int result = 0;
            dt = masterConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string sheetName;
            sheetName = dt.Rows[0]["TABLE_NAME"].ToString();

            sql = "select * from [" + sheetName + "] WHERE [Encoded]<>'YES' OR [Encoded] is NULL ORDER BY [Id]";
            OleDbCommand masterCommand = new OleDbCommand();
            OleDbCommand masterUpdCommand = new OleDbCommand();
            masterUpdCommand.Connection = masterConn;
            masterCommand.CommandText = sql;
            masterCommand.Connection = masterConn;
            OleDbDataReader dr = masterCommand.ExecuteReader();
            
            while (dr.Read())
            {
                if (dr[0].ToString() == "" && dr[1].ToString() == "" && dr[3].ToString() == "" &&
                    dr[4].ToString() == "" && dr[5].ToString() == "" && dr[6].ToString() == "")
                {
                    break;
                }

                searchString = string.Empty;
                if (!string.IsNullOrEmpty(dr["CAMP CODE"].ToString()))
                {
                    masterCampCode = dr["CAMP CODE"].ToString().Trim();
                }
                else
                {
                    masterCampCode = "";
                }
                
                //Position
                if (!string.IsNullOrEmpty(dr["Position"].ToString()))
                {
                    searchString = dr["Position"].ToString().Trim();
                }
                else
                {
                    searchString = searchString + "";
                }
                //Part Number
                if (!string.IsNullOrEmpty(dr["Part Number"].ToString()))
                {
                    searchString = searchString + dr["Part Number"].ToString().Trim();
                }
                else
                {
                    searchString = searchString + "";
                }

                searchString = searchString.Replace("'", "''");
                enCodingCondition = "PP";

                if (masterCampCode != "")
                {                    
                    try
                    {
                        OleDbConnection srcConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @sourceFileName + ";Extended Properties=" + (char)34 + "Excel 8.0;HDR=YES" + (char)34);

                        OleDbCommand srcCommand = new OleDbCommand();

                        srcConn.Open();
                        dt = srcConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        string sheetName1;
                        sheetName1 = dt.Rows[0]["TABLE_NAME"].ToString();

                        srcCommand.Connection = srcConn;

                        string ssql = string.Empty;
                        ssql = "Select Top 1 * FROM [" + sheetName1 + "]  WHERE RTRIM([Position]) + LTRIM([Part Number]) = '" + searchString + "' AND ( [CAMP CODE] is Null OR RTRIM([CAMP CODE]) ='') ORDER BY [Id]";

                        srcCommand.CommandText = ssql;
                        OleDbDataReader selectDr = srcCommand.ExecuteReader();
                        string srcUniqueId = string.Empty;
                        while (selectDr.Read())
                        {
                            srcUniqueId = selectDr["Id"].ToString();
                        }
                        selectDr.Close();
                        if (srcUniqueId != "")
                        {
                            //sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "', [Encoding Condition] ='" + enCodingCondition + "' WHERE RTRIM([Id]) ='" + srcUniqueId + "' ";
                            sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "' WHERE RTRIM([Id]) ='" + srcUniqueId + "' ";
                        }
                        else
                        {
                            //sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "', [Encoding Condition] ='" + enCodingCondition + "' WHERE RTRIM([Position]) + LTRIM([Part Number]) = '" + searchString + "' AND ( [CAMP CODE] is Null OR RTRIM([CAMP CODE]) ='') " ;
                            sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "' WHERE RTRIM([Position]) + LTRIM([Part Number]) = '" + searchString + "' AND ( [CAMP CODE] is Null OR RTRIM([CAMP CODE]) ='') ";
                        }
                        srcCommand.CommandText = sql;
                        result = srcCommand.ExecuteNonQuery();
                                                
                        srcConn.Close();
                        sql = "";
                        if (result > 0)
                        {
                            sql = "Update [" + sheetName + "] SET [Encoded] ='YES' WHERE  [Id]= '" + dr["Id"].ToString().Trim() + "'";
                            masterUpdCommand.CommandText = sql;
                            result = masterUpdCommand.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }

            }
            masterConn.Close();
        }

        public void EncodeSourceFileCampCodesBasedOnPartNumber(string masterFileName, string sourceFileName)
        {
            System.Data.DataTable dt = null;
            OleDbConnection masterConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @masterFileName + ";Extended Properties=" + (char)34 + "Excel 8.0;HDR=Yes" + (char)34);
            masterConn.Open();

            string sql = string.Empty;
            string searchString = string.Empty;
            string masterCampCode = string.Empty;
            string enCodingCondition = string.Empty;
            int result = 0;
            dt = masterConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string sheetName;
            sheetName = dt.Rows[0]["TABLE_NAME"].ToString();

            sql = "select * from [" + sheetName + "] WHERE [Encoded]<>'YES' OR [Encoded] is NULL ORDER BY [Id]";
            OleDbCommand masterCommand = new OleDbCommand();
            OleDbCommand masterUpdCommand = new OleDbCommand();
            masterUpdCommand.Connection = masterConn;
            masterCommand.CommandText = sql;
            masterCommand.Connection = masterConn;
            OleDbDataReader dr = masterCommand.ExecuteReader();

            while (dr.Read())
            {
                if (dr[0].ToString() == "" && dr[1].ToString() == "" && dr[3].ToString() == "" &&
                    dr[4].ToString() == "" && dr[5].ToString() == "" && dr[6].ToString() == "")
                {
                    break;
                }
                searchString = string.Empty;
                if (!string.IsNullOrEmpty(dr["CAMP CODE"].ToString()))
                {
                    masterCampCode = dr["CAMP CODE"].ToString().Trim();
                }
                else
                {
                    masterCampCode = "";
                }
             
                //Part Number
                if (!string.IsNullOrEmpty(dr["Part Number"].ToString()))
                {
                    searchString = dr["Part Number"].ToString().Trim();
                }
                else
                {
                    searchString = searchString + "";
                }

                searchString = searchString.Replace("'", "''");

                enCodingCondition = "PN";

                if (masterCampCode != "")
                {
                    try
                    {
                        OleDbConnection srcConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @sourceFileName + ";Extended Properties=" + (char)34 + "Excel 8.0;HDR=YES" + (char)34);

                        OleDbCommand srcCommand = new OleDbCommand();

                        srcConn.Open();
                        dt = srcConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        string sheetName1;
                        sheetName1 = dt.Rows[0]["TABLE_NAME"].ToString();

                        srcCommand.Connection = srcConn;

                        string ssql = string.Empty;
                        ssql = "Select Top 1 * FROM [" + sheetName1 + "]  WHERE RTRIM([Part Number])  = '" + searchString + "' AND ( [CAMP CODE] is Null OR RTRIM([CAMP CODE]) ='') ORDER BY [Id]";

                        srcCommand.CommandText = ssql;
                        OleDbDataReader selectDr = srcCommand.ExecuteReader();
                        string srcUniqueId = string.Empty;
                        while (selectDr.Read())
                        {
                            srcUniqueId = selectDr["Id"].ToString().Trim();
                        }
                        selectDr.Close();

                        if (srcUniqueId != "")
                        {
                            //sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "', [Encoding Condition] ='" + enCodingCondition + "' WHERE  RTRIM([Id]) ='" + srcUniqueId + "' ";
                            sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "' WHERE  RTRIM([Id]) ='" + srcUniqueId + "' ";
                        }
                        else
                        {
                            //sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "', [Encoding Condition] ='" + enCodingCondition + "' WHERE RTRIM([Part Number]) = '" + searchString + "' AND ( [CAMP CODE] is Null OR RTRIM([CAMP CODE]) ='') ";
                            sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "' WHERE RTRIM([Part Number]) = '" + searchString + "' AND ( [CAMP CODE] is Null OR RTRIM([CAMP CODE]) ='') ";
                        }
                        srcCommand.CommandText = sql;
                        result = srcCommand.ExecuteNonQuery();

                        srcConn.Close();
                        sql = "";
                        if (result > 0)
                        {
                            sql = "Update [" + sheetName + "] SET [Encoded] ='YES' WHERE  [Id]= '" + dr["Id"].ToString().Trim() + "'";
                            masterUpdCommand.CommandText = sql;
                            result = masterUpdCommand.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }

            }
            masterConn.Close();
        }

        public void EncodeSourceFileCampCodesBasedOnDescription(string masterFileName, string sourceFileName)
        {
            System.Data.DataTable dt = null;
            OleDbConnection masterConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @masterFileName + ";Extended Properties=" + (char)34 + "Excel 8.0;HDR=Yes" + (char)34);
            masterConn.Open();

            string sql = string.Empty;
            string searchString = string.Empty;
            string masterCampCode = string.Empty;
            string enCodingCondition = string.Empty;
            int result = 0;
            dt = masterConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string sheetName;
            sheetName = dt.Rows[0]["TABLE_NAME"].ToString();

            sql = "select * from [" + sheetName + "] WHERE [Encoded]<>'YES' OR [Encoded] is NULL ORDER BY [Id]";
            OleDbCommand masterCommand = new OleDbCommand();
            OleDbCommand masterUpdCommand = new OleDbCommand();
            masterUpdCommand.Connection = masterConn;
            masterCommand.CommandText = sql;
            masterCommand.Connection = masterConn;
            OleDbDataReader dr = masterCommand.ExecuteReader();

            while (dr.Read())
            {
                if (dr[0].ToString() == "" && dr[1].ToString() == "" && dr[3].ToString() == "" &&
                    dr[4].ToString() == "" && dr[5].ToString() == "" && dr[6].ToString() == "")
                {
                    break;
                }
                searchString = string.Empty;
                if (!string.IsNullOrEmpty(dr["CAMP CODE"].ToString()))
                {
                    masterCampCode = dr["CAMP CODE"].ToString().Trim();
                }
                else
                {
                    masterCampCode = "";
                }

                //Part Number
                if (!string.IsNullOrEmpty(dr["Description"].ToString()))
                {
                    searchString =  dr["Description"].ToString().Trim();
                }
                else
                {
                    searchString = searchString + "";
                }

                searchString = searchString.Replace("'", "''");

                enCodingCondition = "D";

                if (masterCampCode != "")
                {
                    try
                    {
                        OleDbConnection srcConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @sourceFileName + ";Extended Properties=" + (char)34 + "Excel 8.0;HDR=YES" + (char)34);

                        OleDbCommand srcCommand = new OleDbCommand();

                        srcConn.Open();
                        dt = srcConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        string sheetName1;
                        sheetName1 = dt.Rows[0]["TABLE_NAME"].ToString();

                        srcCommand.Connection = srcConn;

                        string ssql = string.Empty;
                        ssql = "Select Top 1 * FROM [" + sheetName1 + "]  WHERE RTRIM([Description])  = '" + searchString + "' AND ( [CAMP CODE] is Null OR RTRIM([CAMP CODE]) ='') ORDER BY [Id]";

                        srcCommand.CommandText = ssql;
                        OleDbDataReader selectDr = srcCommand.ExecuteReader();
                        string srcUniqueId = string.Empty;
                        while (selectDr.Read())
                        {
                            srcUniqueId = selectDr["Id"].ToString().Trim();
                        }
                        selectDr.Close();

                        if (srcUniqueId != "")
                        {
                            //sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "', [Encoding Condition] ='" + enCodingCondition + "' WHERE  RTRIM([Id]) ='" + srcUniqueId + "' ";
                            sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "' WHERE  RTRIM([Id]) ='" + srcUniqueId + "' ";
                        }
                        else
                        {
                            //sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "', [Encoding Condition] ='" + enCodingCondition + "' WHERE RTRIM([Description]) = '" + searchString + "' AND ( [CAMP CODE] is Null OR RTRIM([CAMP CODE]) ='') ";
                            sql = "Update [" + sheetName1 + "] SET [CAMP CODE] ='" + masterCampCode.ToString().Trim() + "' WHERE RTRIM([Description]) = '" + searchString + "' AND ( [CAMP CODE] is Null OR RTRIM([CAMP CODE]) ='') ";
                        }
                        srcCommand.CommandText = sql;
                        result = srcCommand.ExecuteNonQuery();

                        srcConn.Close();
                        sql = "";
                        if (result > 0)
                        {
                            sql = "Update [" + sheetName + "] SET [Encoded] ='YES' WHERE  [Id]= '" + dr["Id"].ToString().Trim() + "'";
                            masterUpdCommand.CommandText = sql;
                            result = masterUpdCommand.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }

            }
            masterConn.Close();
        }

        public void UpdateSourceDataForDuplicates(string sourceFileName)
        {
            System.Data.DataTable dt = null;            
            OleDbConnection SrcConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @sourceFileName + ";Extended Properties=" + (char)34 + "Excel 8.0;HDR=Yes" + (char)34);
            SrcConn.Open();

            string sql = string.Empty;
            string searchString = string.Empty;
            string CAMPCODE = string.Empty;
            int result = 0;
            dt = SrcConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string sheetName;
            sheetName = dt.Rows[0]["TABLE_NAME"].ToString();

            sql = "SELECT [CAMP CODE], Count(*) AS Count1 from [" + sheetName + "] WHERE [CAMP CODE] is NOT Null AND RTRIM([CAMP CODE]) <> '' GROUP BY [CAMP CODE] HAVING Count(*) >1";
            OleDbCommand srcCommand = new OleDbCommand();
            OleDbCommand srcUpdCommand = new OleDbCommand();
            srcUpdCommand.Connection = SrcConn;
            srcCommand.CommandText = sql;
            srcCommand.Connection = SrcConn;
            OleDbDataReader dr = srcCommand.ExecuteReader();
            
            while (dr.Read())
            {  
                CAMPCODE = dr["CAMP CODE"].ToString().Trim();
                
                searchString = searchString.Replace("'", "''");

                if (CAMPCODE != "" )
                {
                    try
                    {

                        sql = "Update [" + sheetName + "] SET [Duplicate] ='YES' WHERE [CAMP CODE]='" + CAMPCODE + "'";                       
                        srcUpdCommand.CommandText = sql;
                        result = srcUpdCommand.ExecuteNonQuery();
                        
                        sql = "";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }

            }
            SrcConn.Close();
           
        }

        private string validateMasterFile(string masterFileName)
        {
            StringBuilder sb = new StringBuilder();
           
            System.Data.DataTable dt = null;
            OleDbConnection masterConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @masterFileName + ";Extended Properties=" + (char)34 + "Excel 8.0;HDR=Yes" + (char)34);
            masterConn.Open();

            
            dt = masterConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string sheetName;
            sheetName = dt.Rows[0]["TABLE_NAME"].ToString();

            OleDbCommand masterCommand = new OleDbCommand("Select Top 1 * from [" + sheetName + "] ", masterConn);
            OleDbDataAdapter masterAdapter = new OleDbDataAdapter();
            masterAdapter.SelectCommand = masterCommand;
            DataSet masterDataSet = new DataSet();
            masterAdapter.Fill(masterDataSet, "Test");

            masterFileInfo.isIdColumnExists = false;
            masterFileInfo.isCampCodeColumnExists = false;
            masterFileInfo.isEnodedColumnExists = false;
            masterFileInfo.isPositionColumnExists = false;
            masterFileInfo.isDescriptionColumnExists = false;
            masterFileInfo.isPartNumberColumnExists = false;
            masterFileInfo.isValid = false;
            foreach (DataColumn column in masterDataSet.Tables[0].Columns)
            {
                if (column.ColumnName.Trim().ToUpper() == "ID")
                {
                    masterFileInfo.isIdColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "CAMP CODE")
                {
                    masterFileInfo.isCampCodeColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "ENCODED")
                {
                    masterFileInfo.isEnodedColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "POSITION")
                {
                    masterFileInfo.isPositionColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "DESCRIPTION")
                {
                    masterFileInfo.isDescriptionColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "PART NUMBER")
                {
                    masterFileInfo.isPartNumberColumnExists = true;
                }
            }
            if (masterFileInfo.isIdColumnExists && masterFileInfo.isCampCodeColumnExists 
                && masterFileInfo.isEnodedColumnExists && masterFileInfo.isPositionColumnExists 
                && masterFileInfo.isDescriptionColumnExists && masterFileInfo.isPartNumberColumnExists)
            {
                masterFileInfo.isValid = true;
            }

            if (masterFileInfo.isValid)
            {
                sb.Append("");
            }
            else
            {
                sb.AppendLine();
                sb.AppendLine("Master File Error - Following Columns does not Exists in Master File");
                if (masterFileInfo.isIdColumnExists == false)
                {
                    sb.AppendLine("                   -> ID");
                }
                if (masterFileInfo.isCampCodeColumnExists == false)
                {
                    sb.AppendLine("                   -> CAMP CODE ");
                }
                if (masterFileInfo.isEnodedColumnExists == false)
                {
                    sb.AppendLine("                   -> ENCODED ");
                }
                if (masterFileInfo.isPositionColumnExists == false)
                {
                    sb.AppendLine("                   -> POSITION ");
                }
                if (masterFileInfo.isDescriptionColumnExists == false)
                {
                    sb.AppendLine("                   -> DESCRIPTION ");
                }
                if (masterFileInfo.isPartNumberColumnExists == false)
                {
                    sb.AppendLine("                   -> PART NUMBER ");
                }
            }
            masterConn.Close();
            return sb.ToString();
        }
        
        private string validateSourceFile(string sourceFileName)
        {
            StringBuilder sb = new StringBuilder();

            System.Data.DataTable dt = null;
            OleDbConnection sourceConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sourceFileName + ";Extended Properties=" + (char)34 + "Excel 8.0;HDR=Yes" + (char)34);
            sourceConn.Open();


            dt = sourceConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string sheetName;
            sheetName = dt.Rows[0]["TABLE_NAME"].ToString();

            OleDbCommand sourceCommand = new OleDbCommand("Select Top 1 * from [" + sheetName + "] ", sourceConn);
            OleDbDataAdapter sourceAdapter = new OleDbDataAdapter();
            sourceAdapter.SelectCommand = sourceCommand;
            DataSet sourceDataSet = new DataSet();
            sourceAdapter.Fill(sourceDataSet, "Test");

            sourceFileInfo.isIdColumnExists = false;
            sourceFileInfo.isCompCityInstldColumnExists = false;
            sourceFileInfo.isEnodingConditionColumnExists = false;
            sourceFileInfo.isCampCodeColumnExists = false;
            sourceFileInfo.isDuplicateColumnExists = false;
            sourceFileInfo.isAddressedColumnExists = false;
            sourceFileInfo.isPositionColumnExists = false;
            sourceFileInfo.isDescriptionColumnExists = false;
            sourceFileInfo.isPartNumberColumnExists = false;
            sourceFileInfo.isSerialNumberColumnExists = false;
            sourceFileInfo.isPageColumnExists = false;
            sourceFileInfo.isValid = false;

            
            foreach (DataColumn column in sourceDataSet.Tables[0].Columns)
            {
                if (column.ColumnName.Trim().ToUpper() == "ID")
                {
                    sourceFileInfo.isIdColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "COMP_CITY_INSTLD")
                {
                    sourceFileInfo.isCompCityInstldColumnExists = true;
                }
                //if (column.ColumnName.Trim().ToUpper() == "ENCODING CONDITION")
                //{
                //    sourceFileInfo.isEnodingConditionColumnExists = true;
                //}
                if (column.ColumnName.Trim().ToUpper() == "CAMP CODE")
                {
                    sourceFileInfo.isCampCodeColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "DUPLICATE")
                {
                    sourceFileInfo.isDuplicateColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "ADDRESSED")
                {
                    sourceFileInfo.isAddressedColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "POSITION")
                {
                    sourceFileInfo.isPositionColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "DESCRIPTION")
                {
                    sourceFileInfo.isDescriptionColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "PART NUMBER")
                {
                    sourceFileInfo.isPartNumberColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "SERIAL NUMBER")
                {
                    sourceFileInfo.isSerialNumberColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "PAGE")
                {
                    sourceFileInfo.isPageColumnExists = true;
                }
            }
            if (sourceFileInfo.isIdColumnExists && sourceFileInfo.isCompCityInstldColumnExists 
                        && sourceFileInfo.isCampCodeColumnExists 
                        && sourceFileInfo.isDuplicateColumnExists && sourceFileInfo.isAddressedColumnExists 
                        && sourceFileInfo.isPositionColumnExists && sourceFileInfo.isDescriptionColumnExists
                        && sourceFileInfo.isPartNumberColumnExists && sourceFileInfo.isSerialNumberColumnExists && sourceFileInfo.isPageColumnExists)
            {
                sourceFileInfo.isValid = true;
            }
            if (sourceFileInfo.isValid)
            {
                sb.Append("");
            }
            else
            {
                sb.AppendLine();
                sb.AppendLine("Souce File Error - Following Columns does not Exists in Source File");
                if (sourceFileInfo.isIdColumnExists == false)
                {
                    sb.AppendLine("                   -> ID");
                }
                if (sourceFileInfo.isCampCodeColumnExists == false)
                {
                    sb.AppendLine("                   -> CAMP CODE ");
                }
                if (sourceFileInfo.isCompCityInstldColumnExists== false)
                {
                    sb.AppendLine("                   -> COMP_CITY_INSTLD ");
                }
                //if (sourceFileInfo.isEnodingConditionColumnExists == false)
                //{
                //    sb.AppendLine("                   -> ENCODING CONDITION ");
                //}
                if (sourceFileInfo.isDuplicateColumnExists == false)
                {
                    sb.AppendLine("                   -> DUPLICATE ");
                }
                if (sourceFileInfo.isAddressedColumnExists == false)
                {
                    sb.AppendLine("                   -> ADDRESSED ");
                }
                if (sourceFileInfo.isPositionColumnExists == false)
                {
                    sb.AppendLine("                   -> POSITION ");
                }
                if (sourceFileInfo.isDescriptionColumnExists == false)
                {
                    sb.AppendLine("                   -> DESCRIPTION ");
                }
                if (sourceFileInfo.isPartNumberColumnExists == false)
                {
                    sb.AppendLine("                   -> PART NUMBER ");
                }
                if (sourceFileInfo.isSerialNumberColumnExists == false)
                {
                    sb.AppendLine("                   -> SERIAL NUMBER ");
                }
                if (sourceFileInfo.isPageColumnExists == false)
                {
                    sb.AppendLine("                   -> PAGE ");
                }
            }

            sourceConn.Close();
            return sb.ToString();
           
        }

        private string validateProfileFile(string profileFileName)
        {
            StringBuilder sb = new StringBuilder();

            System.Data.DataTable dt = null;
            OleDbConnection profileConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + profileFileName + ";Extended Properties=" + (char)34 + "Excel 8.0;HDR=Yes" + (char)34);
            profileConn.Open();


            dt = profileConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string sheetName;
            sheetName = dt.Rows[0]["TABLE_NAME"].ToString();

            OleDbCommand profileCommand = new OleDbCommand("Select Top 1 * from [" + sheetName + "] ", profileConn);
            OleDbDataAdapter profileAdapter = new OleDbDataAdapter();
            profileAdapter.SelectCommand = profileCommand;
            DataSet profileDataSet = new DataSet();
            profileAdapter.Fill(profileDataSet, "Test");

            profileFileInfo.isIdColumnExists = false;            
            profileFileInfo.isCampCodeColumnExists = false;
            profileFileInfo.isStatusColumnExists = false;
            profileFileInfo.isPnOnColumnExists = false;
            profileFileInfo.isSnOnColumnExists = false;
            profileFileInfo.isWorkPerformedICAOColumnExists = false;
            profileFileInfo.isPageColumnExists = false;
            profileFileInfo.isValid = false;

            foreach (DataColumn column in profileDataSet.Tables[0].Columns)
            {
                if (column.ColumnName.Trim().ToUpper() == "ID")
                {
                    profileFileInfo.isIdColumnExists = true;
                }
               
                if (column.ColumnName.Trim().ToUpper() == "CAMP CODE")
                {
                    profileFileInfo.isCampCodeColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "STATUS")
                {
                    profileFileInfo.isStatusColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "PN ON")
                {
                    profileFileInfo.isPnOnColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "SN ON")
                {
                    profileFileInfo.isSnOnColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "WORK PERFORMED ICAO")
                {
                    profileFileInfo.isWorkPerformedICAOColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "PAGE")
                {
                    profileFileInfo.isPageColumnExists  = true;
                }
               
                
            }
            if (profileFileInfo.isIdColumnExists && profileFileInfo.isCampCodeColumnExists
                        && profileFileInfo.isStatusColumnExists && profileFileInfo.isPnOnColumnExists
                        && profileFileInfo.isSnOnColumnExists && profileFileInfo.isWorkPerformedICAOColumnExists 
                        && profileFileInfo.isPageColumnExists                        )
            {
                profileFileInfo.isValid = true;
            }
            if (profileFileInfo.isValid)
            {
                sb.Append("");
            }
            else
            {
                sb.AppendLine();
                sb.AppendLine("Profile File Error - Following Columns does not Exists in Profile File");
                if (profileFileInfo.isIdColumnExists == false)
                {
                    sb.AppendLine("                   -> ID");
                }
                if (profileFileInfo.isCampCodeColumnExists == false)
                {
                    sb.AppendLine("                   -> CAMP CODE ");
                }
                if (profileFileInfo.isStatusColumnExists == false)
                {
                    sb.AppendLine("                   -> STATUS ");
                }
                if (profileFileInfo.isPnOnColumnExists == false)
                {
                    sb.AppendLine("                   -> PN ON ");
                }
                if (profileFileInfo.isSnOnColumnExists == false)
                {
                    sb.AppendLine("                   -> SN ON ");
                }
                if (profileFileInfo.isWorkPerformedICAOColumnExists == false)
                {
                    sb.AppendLine("                   -> WORK PERFORMED ICAO ");
                }
                if (profileFileInfo.isPageColumnExists == false)
                {
                    sb.AppendLine("                   -> PAGE ");
                }   
            }
            profileConn.Close();
            return sb.ToString();

        }

        private string validateSourceFileForProfileEncoding(string sourceFileName)
        {
            StringBuilder sb = new StringBuilder();

            System.Data.DataTable dt = null;
            OleDbConnection sourceConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sourceFileName + ";Extended Properties=" + (char)34 + "Excel 8.0;HDR=Yes" + (char)34);
            sourceConn.Open();


            dt = sourceConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string sheetName;
            sheetName = dt.Rows[0]["TABLE_NAME"].ToString();

            OleDbCommand sourceCommand = new OleDbCommand("Select Top 1 * from [" + sheetName + "] ", sourceConn);
            OleDbDataAdapter sourceAdapter = new OleDbDataAdapter();
            sourceAdapter.SelectCommand = sourceCommand;
            DataSet sourceDataSet = new DataSet();
            sourceAdapter.Fill(sourceDataSet, "Test");

            sourceFileInfo.isIdColumnExists = false;
            sourceFileInfo.isCompCityInstldColumnExists = false;
            sourceFileInfo.isEnodingConditionColumnExists = false;
            sourceFileInfo.isCampCodeColumnExists = false;
            sourceFileInfo.isDuplicateColumnExists = false;
            sourceFileInfo.isAddressedColumnExists = false;
            sourceFileInfo.isPositionColumnExists = false;
            sourceFileInfo.isDescriptionColumnExists = false;
            sourceFileInfo.isPartNumberColumnExists = false;
            sourceFileInfo.isSerialNumberColumnExists = false;
            sourceFileInfo.isPageColumnExists = false;
            sourceFileInfo.isValid = false;

            foreach (DataColumn column in sourceDataSet.Tables[0].Columns)
            {
                if (column.ColumnName.Trim().ToUpper() == "ID")
                {
                    sourceFileInfo.isIdColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "COMP_CITY_INSTLD")
                {
                    sourceFileInfo.isCompCityInstldColumnExists = true;
                }
                //if (column.ColumnName.Trim().ToUpper() == "ENCODING CONDITION")
                //{
                //    sourceFileInfo.isEnodingConditionColumnExists = true;
                //}
                if (column.ColumnName.Trim().ToUpper() == "CAMP CODE")
                {
                    sourceFileInfo.isCampCodeColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "DUPLICATE")
                {
                    sourceFileInfo.isDuplicateColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "ADDRESSED")
                {
                    sourceFileInfo.isAddressedColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "POSITION")
                {
                    sourceFileInfo.isPositionColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "DESCRIPTION")
                {
                    sourceFileInfo.isDescriptionColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "PART NUMBER")
                {
                    sourceFileInfo.isPartNumberColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "SERIAL NUMBER")
                {
                    sourceFileInfo.isSerialNumberColumnExists = true;
                }
                if (column.ColumnName.Trim().ToUpper() == "PAGE")
                {
                    sourceFileInfo.isPageColumnExists = true;
                }
            }
            //do not check for  Position, Description columns while encoding profile file
            if (sourceFileInfo.isIdColumnExists && sourceFileInfo.isCompCityInstldColumnExists
                        && sourceFileInfo.isCampCodeColumnExists
                        && sourceFileInfo.isDuplicateColumnExists && sourceFileInfo.isAddressedColumnExists
                        && sourceFileInfo.isPartNumberColumnExists && sourceFileInfo.isSerialNumberColumnExists && sourceFileInfo.isPageColumnExists)
            {
                sourceFileInfo.isValid = true;
            }
            if (sourceFileInfo.isValid)
            {
                sb.Append("");
            }
            else
            {
                sb.AppendLine();
                sb.AppendLine("Souce File Error - Following Columns does not Exists in Source File");
                if (sourceFileInfo.isIdColumnExists == false)
                {
                    sb.AppendLine("                   -> ID");
                }
                if (sourceFileInfo.isCampCodeColumnExists == false)
                {
                    sb.AppendLine("                   -> CAMP CODE ");
                }
                if (sourceFileInfo.isCompCityInstldColumnExists == false)
                {
                    sb.AppendLine("                   -> COMP_CITY_INSTLD ");
                }
                
                if (sourceFileInfo.isDuplicateColumnExists == false)
                {
                    sb.AppendLine("                   -> DUPLICATE ");
                }
                if (sourceFileInfo.isAddressedColumnExists == false)
                {
                    sb.AppendLine("                   -> ADDRESSED ");
                }                
                if (sourceFileInfo.isPartNumberColumnExists == false)
                {
                    sb.AppendLine("                   -> PART NUMBER ");
                }
                if (sourceFileInfo.isSerialNumberColumnExists == false)
                {
                    sb.AppendLine("                   -> SERIAL NUMBER ");
                }
                if (sourceFileInfo.isPageColumnExists == false)
                {
                    sb.AppendLine("                   -> PAGE ");
                }
            }

            sourceConn.Close();
            return sb.ToString();

        }
    }
}
