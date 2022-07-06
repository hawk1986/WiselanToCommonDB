/******************************************************************************
 * 2020/02/07 劉燦輝更新
 *  XL員工黃冠智及DT員工李令德的員工編號都是22002，所以在更新時如果沒有加公司，則會錯誤
 * 2020/02/26 劉燦輝更新
 *  將DO及DX的資料排除處理
 ******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace wiselan2GCMIS
{
    class Program
    {
        //將人事資料從智將→GCMIS系統，包含大陸及台北主機
        //轉檔順序 公司→部門→職銜→員工基本資料→員工密碼→員工權限→員工出勤→員工出勤補刷

        private const string DBUSER = "MISUser";
        private const string DBPWD = "1234QWERasdf";
        //private const string strGCMIS = "Data Source=twtpe1wsm18;User ID=" + DBUSER + ";Password=" + DBPWD + ";Initial Catalog=GCMIS";
        //private const string strCommon_L = "Data Source=twtpe1wsm18;User ID=" + DBUSER + ";Password=" + DBPWD + ";Initial Catalog=CommonDB";
        //private const string strWiselan = "Data Source=twtpe1sqm09;User ID=hr;Password=wiselan;Initial Catalog=WAEGIS_SQL";
        //private const string strHai = "Data Source=twtpe1sqm18;User ID=" + DBUSER + ";Password=" + DBPWD + ";Initial Catalog=Hai;pooling=false";

        #region 在 Server上使用的連結字串

        private const string strGCMIS = "Data Source=twtpe1sqm01;User ID=" + DBUSER + ";Password=" + DBPWD + ";Initial Catalog=GCMIS;pooling=false";
        private const string strCommon_L = "Data Source=twtpe1sqm01;User ID=" + DBUSER + ";Password=" + DBPWD + ";Initial Catalog=CommonDB;pooling=false";
        private const string strWiselan = "Data Source=twtpe1sqm01;User ID=hr;Password=wiselan;Initial Catalog=WAEGIS_SQL;pooling=false";
        private const string strHai = "Data Source=twtpe1sqm01;User ID=" + DBUSER + ";Password=" + DBPWD + ";Initial Catalog=Hai;pooling=false";
        #endregion

        #region 在測試機上的連結字串
        //private const string strGCMIS = "Data Source=twtpe1sqm05;User ID=" + DBUSER + ";Password=" + DBPWD + ";Initial Catalog=GCMIS";
        //private const string strCommon_L = "Data Source=twtpe1sqm05;User ID=" + DBUSER + ";Password=" + DBPWD + ";Initial Catalog=CommonDB";
        //private const string strWiselan = "Data Source=twtpe1sqm09;User ID=hr;Password=wiselan;Initial Catalog=WAEGIS_SQL";
        #endregion

        #region 全體變數
        private const string PROG_ID = "wiselan2GCMIS";
        private const string PROG_USER_ID = "Schedule Task";
        private const string RECIPIENT = "robert.liu@dentsuaegis.com";
        static DataTable dtBrand = new DataTable();
        #endregion

        static void Main(string[] args)
        {
            #region 正式版
            bool bCompany = false;      //轉公司資料
            bool bDept = true;         //轉部門資料
            bool bTitle = true;        //轉職銜資料
            bool bEmployee = true;     //轉員工基本資料
            bool bTran = true;         //轉員工資料異動
            bool bBoss = true;         //轉員工主管資料
            bool bExt = true;          //轉分機資料
            bool bEar = true;          //轉耳機資料
            bool bAttendance = true;   //轉員工出勤資料
            bool bFill = true;         //轉員工出勤補刷資料
            #endregion

            #region 測試時用
            //bool bCompany = false;      //轉公司資料
            //bool bDept = false;         //轉部門資料
            //bool bTitle = false;        //轉職銜資料
            //bool bEmployee = false;     //轉員工基本資料
            //bool bTran = false;         //轉員工資料異動
            //bool bBoss = true;         //轉員工主管資料
            //bool bExt = false;          //轉分機資料
            //bool bEar = false;           //轉耳機資料
            //bool bAttendance = false;   //轉員工出勤資料
            //bool bFill = false;          //轉員工出勤補刷資料
            #endregion


            string sComputerName = System.Environment.GetEnvironmentVariable("COMPUTERNAME").ToString();
            string sDBType = "";
            string sActive = "";
            string sComp = "";
            string sDept = "";
            string sTitle = "";
            

            #region 轉公司資料
            if (bCompany)
            {
                //using (SqlConnection cnWL = new SqlConnection(strWiselan))
                //{
                //    try
                //    {
                //        cnWL.Open();
                //        SqlCommand cmdWL = new SqlCommand("SELECT comno, comnm, comnm2 FROM company", cnWL);
                //        SqlDataReader rdWL = cmdWL.ExecuteReader();
                //        if (rdWL.HasRows)
                //        {
                //            SqlConnection cnCmnR = new SqlConnection(strCommon_R);
                //            SqlConnection cnCmnL = new SqlConnection(strCommon_L);
                //            cnCmnR.Open();
                //            cnCmnL.Open();

                //            SqlCommand cmdR = new SqlCommand();
                //            SqlCommand cmdL = new SqlCommand();
                //            cmdR.Connection = cnCmnR;
                //            cmdL.Connection = cnCmnL;

                //            SqlDataReader rdR;
                //            SqlDataReader rdL;

                //            while (rdWL.Read())
                //            {
                //                cmdR.CommandText = "SELECT CompCode FROM Company WHERE RegionCode = @region AND CompCode = @Comp";
                //                cmdR.Parameters.Clear();
                //                cmdR.Parameters.Add(new SqlParameter("@region", SqlDbType.VarChar));
                //                cmdR.Parameters.Add(new SqlParameter("@comp", SqlDbType.VarChar));
                //                cmdR.Parameters["@region"].Value = "TW";
                //                cmdR.Parameters["@comp"].Value = rdWL.GetString(0).Trim();
                //                rdR = cmdR.ExecuteReader();
                //                if (!rdR.HasRows)
                //                {
                //                    SqlCommand cmdR1 = new SqlCommand();
                //                    cmdR1.Connection = cnCmnR;
                //                    cmdR1.CommandText = "INSERT INTO Company (RegionCode, CompCode, CompEngShortTitle, " +
                //                        "CompEngFullTitle, CompLocShortTitle, CompLocFullTitle) VALUES (@region, @comp, " +
                //                        "@engshort, @engfull, @locshort, @locfull)";
                //                    cmdR1.Parameters.Clear();
                //                    cmdR1.Parameters.Add(new SqlParameter("@region", SqlDbType.VarChar));
                //                    cmdR1.Parameters.Add(new SqlParameter("@comp", SqlDbType.VarChar));
                //                    cmdR1.Parameters.Add(new SqlParameter("@engshort", SqlDbType.VarChar));
                //                    cmdR1.Parameters.Add(new SqlParameter("@engfull", SqlDbType.VarChar));
                //                    cmdR1.Parameters.Add(new SqlParameter("@locshort", SqlDbType.NVarChar));
                //                    cmdR1.Parameters.Add(new SqlParameter("@locfull", SqlDbType.NVarChar));
                //                    cmdR1.Parameters["@region"].Value = "TW";
                //                    cmdR1.Parameters["@comp"].Value = rdWL.GetString(0).Trim();
                //                    string[] sEng = rdWL.GetString(2).Trim().Split(' ');
                //                    cmdR1.Parameters["@engshort"].Value = sEng[0];
                //                    cmdR1.Parameters["@engfull"].Value = rdWL.GetString(2).Trim();
                //                    cmdR1.Parameters["@locshort"].Value = rdWL.GetString(1).Trim();
                //                    cmdR1.Parameters["@locfull"].Value = rdWL.GetString(1).Trim();
                //                    cmdR1.ExecuteNonQuery();
                //                }
                //                rdR.Close();

                //                cmdL.CommandText = "SELECT CompCode FROM Company WHERE RegionCode = @region AND CompCode = @Comp";
                //                cmdL.Parameters.Clear();
                //                cmdL.Parameters.Add(new SqlParameter("@region", SqlDbType.VarChar));
                //                cmdL.Parameters.Add(new SqlParameter("@comp", SqlDbType.VarChar));
                //                cmdL.Parameters["@region"].Value = "TW";
                //                cmdL.Parameters["@comp"].Value = rdWL.GetString(0).Trim();
                //                rdL = cmdL.ExecuteReader();
                //                if (!rdL.HasRows)
                //                {
                //                    SqlCommand cmdL1 = new SqlCommand();
                //                    cmdL1.Connection = cnCmnL;
                //                    cmdL1.CommandText = "INSERT INTO Company (RegionCode, CompCode, CompEngShortTitle, " +
                //                        "CompEngFullTitle, CompLocShortTitle, CompLocFullTitle) VALUES (@region, @comp, " +
                //                        "@engshort, @engfull, @locshort, @locfull)";
                //                    cmdL1.Parameters.Clear();
                //                    cmdL1.Parameters.Add(new SqlParameter("@region", SqlDbType.VarChar));
                //                    cmdL1.Parameters.Add(new SqlParameter("@comp", SqlDbType.VarChar));
                //                    cmdL1.Parameters.Add(new SqlParameter("@engshort", SqlDbType.VarChar));
                //                    cmdL1.Parameters.Add(new SqlParameter("@engfull", SqlDbType.VarChar));
                //                    cmdL1.Parameters.Add(new SqlParameter("@locshort", SqlDbType.NVarChar));
                //                    cmdL1.Parameters.Add(new SqlParameter("@locfull", SqlDbType.NVarChar));
                //                    cmdL1.Parameters["@region"].Value = "TW";
                //                    cmdL1.Parameters["@comp"].Value = rdWL.GetString(0).Trim();
                //                    string[] sEng = rdWL.GetString(2).Trim().Split(' ');
                //                    cmdL1.Parameters["@engshort"].Value = sEng[0];
                //                    cmdL1.Parameters["@engfull"].Value = rdWL.GetString(2).Trim();
                //                    cmdL1.Parameters["@locshort"].Value = rdWL.GetString(1).Trim();
                //                    cmdL1.Parameters["@locfull"].Value = rdWL.GetString(1).Trim();
                //                    cmdL1.ExecuteNonQuery();
                //                }
                //                rdL.Close();
                //            }
                //            cmdR.Dispose();
                //            cmdL.Dispose();
                //        }
                //    }
                //    catch (Exception ex)
                //    {
                //        DANCom.DANClass.WriteToLog(PROG_ID, 0, "E", PROG_USER_ID, sComputerName, ex.Message);

                //    }

                //}
            }
            #endregion

            #region 轉部門資料
            if (bDept)
            {
                //DANCom.DANClass.WriteToLog(PROG_ID, 1, "D", PROG_USER_ID, sComputerName, "開始轉部門資料");

                sDBType = "Wiselan";
                using (SqlConnection cnWLComp = new SqlConnection(strWiselan))
                {
                    try
                    {
                        cnWLComp.Open();

                        SqlCommand cmdWLComp = new SqlCommand();
                        cmdWLComp.Connection = cnWLComp;
                        cmdWLComp.CommandText = "SELECT comno FROM Company";
                        SqlDataReader rdWLComp = cmdWLComp.ExecuteReader();     //取得智將的公司資料
                        if (rdWLComp.HasRows)
                        {
                            SqlConnection cnL = new SqlConnection(strCommon_L);
                            SqlCommand cmdL = new SqlCommand();
                            SqlCommand cmdL_Write = new SqlCommand();
                            SqlCommand cmdWL = new SqlCommand();
                            SqlConnection cnWL = new SqlConnection(strWiselan);
                            cmdWL.Connection = cnWL;
                            cnWL.Open();
                            sDBType = "Local";
                            cnL.Open();
                            cmdL.Connection = cnL;
                            cmdL_Write.Connection = cnL;

                            while (rdWLComp.Read())
                            {
                                cmdWL.CommandText = "SELECT depno, depnm, edepnm FROM Depmsf";
                                SqlDataReader rdWL = cmdWL.ExecuteReader(); //取得智將的部門資料
                                if (rdWL.HasRows)
                                {
                                    while (rdWL.Read())
                                    {
                                        //如果是貝立德或台灣電通的資料則跳過
                                        if (rdWLComp.GetString(0).Trim() == "DX" || rdWLComp.GetString(0).Trim() == "DT") continue;

                                        sDBType = "Local";
                                        sActive = "Read";
                                        cmdL.CommandText = "SELECT CompCode, DeptCode, ISNULL(DeptEngFullName, ''), ISNULL(DeptLocFullName, '') FROM Dept WHERE RegionCode = 'TW' AND CompCode = '" + rdWLComp.GetString(0).Trim() + "' AND DeptCode = '" + rdWL.GetString(0).Trim() + "'";
                                        SqlDataReader rdL = cmdL.ExecuteReader();   //取得GCMIS本地部門資料
                                        if (rdL.HasRows)
                                        {
                                            rdL.Read();
                                            if ((rdWL.GetString(2).Trim() != rdL.GetString(2).Trim()) || (rdWL.GetString(1).Trim() != rdL.GetString(3).Trim()))
                                            {
                                                rdL.Close();
                                                sActive = "Update";
                                                cmdL_Write.CommandText = "UPDATE Dept SET DeptEngShortName = @eng, DeptEngFullName = @eng, DeptLocShortName = @loc, DeptLocFullName = @loc WHERE RegionCode = 'TW' AND CompCode = @comp AND DeptCode = @dept";
                                                cmdL_Write.Parameters.Clear();
                                                cmdL_Write.Parameters.Add(new SqlParameter("@eng", SqlDbType.VarChar));
                                                cmdL_Write.Parameters.Add(new SqlParameter("@loc", SqlDbType.NVarChar));
                                                cmdL_Write.Parameters.Add(new SqlParameter("@comp", SqlDbType.VarChar));
                                                cmdL_Write.Parameters.Add(new SqlParameter("@dept", SqlDbType.VarChar));
                                                cmdL_Write.Parameters["@eng"].Value = rdWL.GetString(2).Trim();
                                                cmdL_Write.Parameters["@loc"].Value = rdWL.GetString(1).Trim();
                                                cmdL_Write.Parameters["@comp"].Value = rdWLComp.GetString(0).Trim();
                                                cmdL_Write.Parameters["@dept"].Value = rdWL.GetString(0).Trim();
                                                cmdL_Write.ExecuteNonQuery();       //更新GCMIS本地部門資料
                                            }
                                            else
                                                rdL.Close();
                                        }
                                        else
                                        {
                                            rdL.Close();
                                            sActive = "Insert";
                                            cmdL_Write.CommandText = "INSERT INTO Dept (RegionCode, CompCode, DeptCode, DeptEngShortName, DeptEngFullName, DeptLocShortName, DeptLocFullName, CloseFlag) " +
                                                " VALUES ('TW', @comp, @dept, @deptengshort, @deptengfull, @deptlocshort, @deptlocfull, 'N')";
                                            cmdL_Write.Parameters.Clear();
                                            cmdL_Write.Parameters.Add(new SqlParameter("@comp", SqlDbType.VarChar));
                                            cmdL_Write.Parameters.Add(new SqlParameter("@dept", SqlDbType.VarChar));
                                            cmdL_Write.Parameters.Add(new SqlParameter("@deptengshort", SqlDbType.VarChar));
                                            cmdL_Write.Parameters.Add(new SqlParameter("@deptengfull", SqlDbType.VarChar));
                                            cmdL_Write.Parameters.Add(new SqlParameter("@deptlocshort", SqlDbType.NVarChar));
                                            cmdL_Write.Parameters.Add(new SqlParameter("@deptlocfull", SqlDbType.NVarChar));
                                            cmdL_Write.Parameters["@comp"].Value = rdWLComp.GetString(0).Trim();
                                            cmdL_Write.Parameters["@dept"].Value = rdWL.GetString(0).Trim();
                                            cmdL_Write.Parameters["@deptengshort"].Value = rdWL.GetString(2).Trim();
                                            cmdL_Write.Parameters["@deptengfull"].Value = rdWL.GetString(2).Trim();
                                            cmdL_Write.Parameters["@deptlocshort"].Value = rdWL.GetString(1).Trim().Length > 20 ? rdWL.GetString(1).Trim().Substring(0, 20) : rdWL.GetString(1).Trim();
                                            cmdL_Write.Parameters["@deptlocfull"].Value = rdWL.GetString(1).Trim();
                                            cmdL_Write.ExecuteNonQuery();           //新增GCMISd本地部門資料
                                        }
                                    }
                                }
                                rdWL.Close();
                            }

                        }
                        rdWLComp.Close();
                        //DANCom.DANClass.WriteToLog(PROG_ID, 1, "D", PROG_USER_ID, sComputerName, "部門資料轉檔結束");
                    }
                    catch (Exception ex)
                    {
                        DANCom.DANClass.SendMail(RECIPIENT, "部門資料轉檔錯誤！", sDBType + " - " + sActive + "公司：" + sComp + "-部門：" + sDept + " - 部門資料轉檔錯誤，請檢查！");
                        DANCom.DANClass.WriteToLog(PROG_ID, 1, "E", PROG_USER_ID, sComputerName, sDBType + " - " + ex.Message);
                    }
                }
            }
            #endregion

            #region 轉職銜資料
            if (bTitle)
            {
                //DANCom.DANClass.WriteToLog(PROG_ID, 2, "D", PROG_USER_ID, sComputerName, "開始轉職銜資料");

                sDBType = "Wiselan";
                using (SqlConnection cnWLComp = new SqlConnection(strWiselan))
                {
                    try
                    {
                        cnWLComp.Open();

                        SqlCommand cmdWLComp = new SqlCommand();
                        cmdWLComp.Connection = cnWLComp;
                        cmdWLComp.CommandText = "SELECT comno FROM Company";
                        SqlDataReader rdWLComp = cmdWLComp.ExecuteReader(); //取得智將公司資料
                        if (rdWLComp.HasRows)
                        {
                            //SqlConnection cnR = new SqlConnection(strCommon_R);
                            SqlConnection cnL = new SqlConnection(strCommon_L);
                            //SqlCommand cmdR = new SqlCommand();
                            SqlCommand cmdL = new SqlCommand();
                            //SqlCommand cmdR_Write = new SqlCommand();
                            SqlCommand cmdL_Write = new SqlCommand();
                            SqlCommand cmdWL = new SqlCommand();
                            SqlConnection cnWL = new SqlConnection(strWiselan);
                            cmdWL.Connection = cnWL;
                            cnWL.Open();
                            //sDBType = "Remote";
                            //cnR.Open();
                            sDBType = "Local";
                            cnL.Open();
                            //cmdR.Connection = cnR;
                            //cmdR_Write.Connection = cnR;
                            cmdL.Connection = cnL;
                            cmdL_Write.Connection = cnL;

                            while (rdWLComp.Read())
                            {
                                cmdWL.CommandText = "SELECT tilno, tilnm, etilnm FROM Tilmsf";
                                SqlDataReader rdWL = cmdWL.ExecuteReader(); //取得智將頭銜資料
                                if (rdWL.HasRows)
                                {
                                    while (rdWL.Read())
                                    {
                                        //如果是貝立德或台灣電通則跳過
                                        if (rdWLComp.GetString(0).Trim() == "DX" || rdWLComp.GetString(0).Trim() == "DT") continue;

                                        //sDBType = "Remote";
                                        //sActive = "Read";
                                        sComp = rdWLComp.GetString(0);
                                        sTitle = rdWL.GetString(0);
                                        //cmdR.CommandText = "SELECT CompCode, JobTitleCode, JobEngTitle, JobLocTitle FROM JobTitle WHERE RegionCode = 'TW' AND CompCode = '" + rdWLComp.GetString(0).Trim() + "' AND JobTitleCode = '" + rdWL.GetString(0).Trim() + "'";
                                        //SqlDataReader rdR = cmdR.ExecuteReader();   //取得GCMIS遠端頭銜資料
                                        //if (rdR.HasRows)
                                        //{
                                        //    rdR.Read();
                                        //    if ((rdWL.GetString(2).Trim() != rdR.GetString(2).Trim()) || (rdWL.GetString(1).Trim() != rdR.GetString(3).Trim()))
                                        //    {
                                        //        rdR.Close();
                                        //        sActive = "Update";
                                        //        cmdR_Write.CommandText = "UPDATE JobTitle SET JobEngTitle = '" + rdWL.GetString(2).Trim() + "', JobLocTitle = '" + rdWL.GetString(1).Trim() + "' WHERE RegionCode = 'TW' AND CompCode = '" + rdWLComp.GetString(0).Trim() + "' AND JobTitleCode = '" + rdWL.GetString(0).Trim() + "'";
                                        //        cmdR_Write.ExecuteNonQuery();       //更新GCMIS遠端頭銜資料
                                        //    }
                                        //    else
                                        //        rdR.Close();
                                        //}
                                        //else
                                        //{
                                        //    rdR.Close();
                                        //    sActive = "Insert";
                                        //    cmdR_Write.CommandText = "INSERT INTO JobTitle (RegionCode, CompCode, JobTitleCode, JobEngTitle, JobLocTitle) " +
                                        //        " VALUES ('TW', '" + rdWLComp.GetString(0).Trim() + "', '" + rdWL.GetString(0).Trim() + "', '" + rdWL.GetString(2).Trim() + "','" + rdWL.GetString(1).Trim() + "')";
                                        //    cmdR_Write.ExecuteNonQuery();           //新增GCMIS遠端頭銜資料
                                        //}

                                        sDBType = "Local";
                                        sActive = "Read";
                                        cmdL.CommandText = "SELECT CompCode, JobTitleCode, JobEngTitle, JobLocTitle FROM JobTitle WHERE RegionCode = 'TW' AND CompCode = '" + rdWLComp.GetString(0).Trim() + "' AND JobTitleCode = '" + rdWL.GetString(0).Trim() + "'";
                                        SqlDataReader rdL = cmdL.ExecuteReader();
                                        if (rdL.HasRows)
                                        {
                                            rdL.Read();
                                            if ((rdWL.GetString(2).Trim() != rdL.GetString(2).Trim()) || (rdWL.GetString(1).Trim() != rdL.GetString(3).Trim()))
                                            {
                                                rdL.Close();
                                                sActive = "Update";
                                                cmdL_Write.CommandText = "UPDATE JobTitle SET JobEngTitle = '" + rdWL.GetString(2).Trim() + "', JobLocTitle = '" + rdWL.GetString(1).Trim() + "' WHERE RegionCode = 'TW' AND CompCode = '" + rdWLComp.GetString(0).Trim() + "' AND JobTitleCode = '" + rdWL.GetString(0).Trim() + "'";
                                                cmdL_Write.ExecuteNonQuery();       //更新GCMIS本地頭銜資料
                                            }
                                            else
                                                rdL.Close();
                                        }
                                        else
                                        {
                                            rdL.Close();
                                            sActive = "Insert";
                                            cmdL_Write.CommandText = "INSERT INTO JobTitle (RegionCode, CompCode, JobTitleCode, JobEngTitle, JobLocTitle) " +
                                                " VALUES ('TW', '" + rdWLComp.GetString(0).Trim() + "', '" + rdWL.GetString(0).Trim() + "', '" + rdWL.GetString(2).Trim() + "','" + rdWL.GetString(1).Trim() + "')";
                                            cmdL_Write.ExecuteNonQuery();           //新增GCMIS本地頭銜資料
                                        }
                                    }
                                }
                                rdWL.Close();
                            }

                        }
                        rdWLComp.Close();
                        //DANCom.DANClass.WriteToLog(PROG_ID, 2, "D", PROG_USER_ID, sComputerName, "職銜資料轉檔結束");
                    }
                    catch (Exception ex)
                    {
                        DANCom.DANClass.SendMail(RECIPIENT, "職銜資料轉檔錯誤！", sDBType + " - " + sActive + " - 公司：" + sComp + "-職銜：" + sTitle + " - 職銜資料轉檔錯誤，請檢查！");
                        DANCom.DANClass.WriteToLog(PROG_ID, 2, "E", PROG_USER_ID, sComputerName, sDBType + " - " + ex.Message);
                    }
                }
            }
            #endregion

            #region 轉員工資料
            if (bEmployee)
            {
                //DANCom.DANClass.WriteToLog(PROG_ID, 3, "D", PROG_USER_ID, sComputerName, "開始轉員工基本資料");
                string sStatment = "";
                string sEmpNo = "";
                //int iUniqNo = 0;

                try
                {
                    sDBType = "Wiselan";
                    using (SqlConnection cnWL = new SqlConnection(strWiselan))
                    {
                        //先抓出在日期範圍有異動的員工編號
                        cnWL.Open();
                        
                        SqlDataAdapter adaBrand = new SqlDataAdapter("SELECT brandno, brandnm FROM Brandtab", cnWL);
                        adaBrand.Fill(dtBrand);

                        SqlCommand cmdWL = new SqlCommand("SELECT Distinct(empno) FROM Trace WHERE tables = 'Empmsf' AND chdatetime >= @chg", cnWL);
                        cmdWL.Parameters.Clear();
                        cmdWL.Parameters.Add(new SqlParameter("@chg", SqlDbType.VarChar));
                        cmdWL.Parameters["@chg"].Value = DateTime.Today.AddMonths(-1).ToString("yyyy/MM/dd");
                        SqlDataReader rdWL = cmdWL.ExecuteReader();
                        if (rdWL.HasRows)
                        {
                            SqlConnection cnL = new SqlConnection(strCommon_L);
                            sDBType = "Local";
                            cnL.Open();
                            SqlCommand cmdL = new SqlCommand();
                            cmdL.Connection = cnL;
                            SqlDataReader rdL;

                            while (rdWL.Read())
                            {
                                //int iUniqNo = 0;
                                sEmpNo = rdWL.GetString(0).Trim();

                                if (sEmpNo == "10018")
                                    sDBType = "text";
                                using (SqlConnection cnW = new SqlConnection(strWiselan))
                                {
                                    sDBType = "Wiselan-A";
                                    cnW.Open();
                                    //                                         0     1       2      3     4      5     6       7     8     9      10      11     12        13    14     15       16
                                    SqlCommand cmdW = new SqlCommand("SELECT empno, empnm, empnm2, wdat, odat, depno, tilno, email, tel2, tel3, topman, comno, directno, tel_no, sex, hris_id, brandno FROM Empmsf WHERE empno = @no", cnW);
                                    cmdW.Parameters.Clear();
                                    cmdW.Parameters.Add(new SqlParameter("@no", SqlDbType.Char));
                                    cmdW.Parameters["@no"].Value = rdWL.GetString(0);
                                    SqlDataReader rdW = cmdW.ExecuteReader();
                                    if (rdW.HasRows)
                                    {
                                        rdW.Read();

                                        //檢查如果是貝立德或台灣電通的員工就跳過
                                        if (rdW.GetString(11).CompareTo("DX") == 0 || rdW.GetString(11).CompareTo("DT") == 0)
                                        {
                                            rdW.Close();
                                            continue;
                                        }

                                        sEmpNo = rdW.GetString(0);
                                        if (rdW.GetString(4).Trim() == "")
                                        {
                                            // 以員工編號查詢該員工資料是否已存在 0           1           2            3             4          5         6               7            8               9              10          11            12             13        14     15        16
                                            cmdL.CommandText = "SELECT EmpUniqNo, EmpLocName, EmpEngName, EmpStartDate, EmpQuitDate, DeptCode, JobTitleCode, EmpCompEmail, EmpHomePhone, EmpMobilePhone, BossEmpUniqNo, CompCode, NextBossEmpUniqNo, EmpExtNo, EmpGender, WDID, BrandCode, EmpNo FROM Employee WHERE RegionCode = 'TW' AND (CompCode <> 'DT' AND CompCode <> 'DX') AND (EmpQuitDate IS NULL OR EmpQuitDate = '' OR EmpQuitDate >= (GetDate() - 1)) AND EmpNo = @no AND CompCode = @comp";
                                            //cmdL.CommandText = "SELECT EmpUniqNo, EmpLocName, EmpEngName, EmpStartDate, EmpQuitDate, DeptCode, JobTitleCode, EmpCompEmail, EmpHomePhone, EmpMobilePhone, BossEmpUniqNo, CompCode, NextBossEmpUniqNo, EmpExtNo, EmpGender, WDID, EmpAdLoginID FROM Employee WHERE RegionCode = 'TW' AND (EmpQuitDate IS NULL OR EmpQuitDate = '' OR EmpQuitDate >= GetDate()) AND EmpNo = @no";
                                        }
                                        else
                                        {
                                            cmdL.CommandText = "SELECT EmpUniqNo, EmpLocName, EmpEngName, EmpStartDate, EmpQuitDate, DeptCode, JobTitleCode, EmpCompEmail, EmpHomePhone, EmpMobilePhone, BossEmpUniqNo, CompCode, NextBossEmpUniqNo, EmpExtNo, EmpGender, WDID, BrandCode, EmpNo FROM Employee WHERE RegionCode = 'TW' AND (CompCode <> 'DT' AND CompCode <> 'DX') AND EmpNo = @no AND CompCode = @comp";
                                        }
                                        cmdL.Parameters.Clear();
                                        cmdL.Parameters.Add(new SqlParameter("@no", SqlDbType.VarChar));
                                        cmdL.Parameters.Add(new SqlParameter("@comp", SqlDbType.VarChar));
                                        cmdL.Parameters["@no"].Value = rdW.GetString(0).Trim();
                                        cmdL.Parameters["@comp"].Value = rdW.GetString(11).Trim();
                                        rdL = cmdL.ExecuteReader();
                                        if (rdL.HasRows)        // 員工編號已存在
                                        {
                                            #region 員工編號已存在
                                            rdL.Read();
                                            //if (rdW.GetString(1).Trim() != rdL.GetString(1) ||
                                            //    rdW.GetString(2).Trim() != rdL.GetString(2) ||
                                            //    rdW.GetString(3).Trim() != rdL.GetDateTime(3).ToString("yyyy/MM/dd") ||
                                            //    rdW.GetString(4).Trim() != (rdL.IsDBNull(4) ? "" : rdL.GetDateTime(4).ToString("yyyy/MM/dd")) ||
                                            //    rdW.GetString(5).Trim() != (rdL.IsDBNull(5) ? "" : rdL.GetString(5)) ||
                                            //    rdW.GetString(6).Trim() != (rdL.IsDBNull(6) ? "" : rdL.GetString(6)) ||
                                            //    rdW.GetString(7).Trim() != (rdL.IsDBNull(7) ? "" : rdL.GetString(7)) ||
                                            //    rdW.GetString(8).Trim() != (rdL.IsDBNull(8) ? "" : rdL.GetString(8)) ||
                                            //    rdW.GetString(9).Trim() != (rdL.IsDBNull(9) ? "" : rdL.GetString(9)) ||
                                            //    rdW.GetString(10).Trim() != GetEmpNo((rdL.IsDBNull(10) ? 0 : rdL.GetInt32(10)), strCommon_L) ||
                                            //    (rdW.GetString(11).Trim() != (rdL.IsDBNull(11) ? "" : rdL.GetString(11)) && ((rdW.GetString(11).Trim() != "PO") || (rdL.GetString(11).Trim() != "AMP"))) ||
                                            //    rdW.GetString(12).Trim() != GetEmpNo((rdL.IsDBNull(12) ? 0 : rdL.GetInt32(12)), strCommon_L) ||
                                            //    rdW.GetString(13).Trim() != (rdL.IsDBNull(13) ? "" : rdL.GetString(13)) ||
                                            //    (rdW.GetString(14).Trim() == "1" ? "M" : "F") != (rdL.IsDBNull(14) ? "" : rdL.GetString(14)) ||
                                            //    rdW.GetString(15).Trim() != (rdL.IsDBNull(15) ? "" : rdL.GetString(15)) || 
                                            //    rdW.GetString(16).Trim() != (rdL.IsDBNull(16) ? "" : rdL.GetString(16)))
                                            if (rdW.GetString(1).Trim() != rdL.GetString(1) ||
                                                rdW.GetString(2).Trim() != rdL.GetString(2) ||
                                                rdW.GetString(3).Trim() != rdL.GetDateTime(3).ToString("yyyy/MM/dd") ||
                                                rdW.GetString(4).Trim() != (rdL.IsDBNull(4) ? "" : rdL.GetDateTime(4).ToString("yyyy/MM/dd")) ||
                                                rdW.GetString(5).Trim() != (rdL.IsDBNull(5) ? "" : rdL.GetString(5)) ||
                                                rdW.GetString(6).Trim() != (rdL.IsDBNull(6) ? "" : rdL.GetString(6)) ||
                                                rdW.GetString(7).Trim() != (rdL.IsDBNull(7) ? "" : rdL.GetString(7)) ||
                                                rdW.GetString(8).Trim() != (rdL.IsDBNull(8) ? "" : rdL.GetString(8)) ||
                                                rdW.GetString(9).Trim() != (rdL.IsDBNull(9) ? "" : rdL.GetString(9)) ||
                                                rdW.GetString(10).Trim() != GetEmpNo((rdL.IsDBNull(10) ? 0 : rdL.GetInt32(10)), strCommon_L) ||
                                                (rdW.GetString(11).Trim() != (rdL.IsDBNull(11) ? "" : rdL.GetString(11)) && ((rdW.GetString(11).Trim() != "PO") || (rdL.GetString(11).Trim() != "AMP"))) ||
                                                rdW.GetString(13).Trim() != (rdL.IsDBNull(13) ? "" : rdL.GetString(13)) ||
                                                (rdW.GetString(14).Trim() == "1" ? "M" : "F") != (rdL.IsDBNull(14) ? "" : rdL.GetString(14)) ||
                                                rdW.GetString(15).Trim() != (rdL.IsDBNull(15) ? "" : rdL.GetString(15)) ||
                                                GetBrandName(rdW.GetString(16).Trim(), rdW.GetString(11).Trim(), rdW.GetString(5).Trim()) != (rdL.IsDBNull(16) ? "" : rdL.GetString(16)))
                                            {
                                                #region 若資料有異動則更新
                                                int iUniq = rdL.GetInt32(0);
                                                rdL.Close();

                                                cmdL.CommandText = "UPDATE Employee SET EmpLocName = @locname, EmpEngName = @engname, " +
                                                    "EmpStartDate = @startdate, EmpQuitDate = @quitdate, DeptCode = @dept, CompCode = @comp, " +
                                                    "JobTitleCode = @job, EmpCompEmail = @email, EmpHomePhone = @home, EmpMobilePhone = @mobile, " +
                                                    "BossEmpUniqNo = @boss, NextBossEmpUniqNo = @next, RegionCode = 'TW', WDID = @wdid, " +
                                                    "EmpGender = @gender, BrandCode = @brand, ChangeBe = @change " +
                                                    " WHERE EmpUniqNo = @no";
                                                //cmdL.CommandText = "UPDATE Employee SET EmpLocName = @locname, EmpEngName = @engname, " +
                                                //    "EmpStartDate = @startdate, EmpQuitDate = @quitdate, DeptCode = @dept, CompCode = @comp, " +
                                                //    "JobTitleCode = @job, EmpCompEmail = @email, EmpHomePhone = @home, EmpMobilePhone = @mobile, " +
                                                //    "BossEmpUniqNo = @boss, NextBossEmpUniqNo = @next, RegionCode = 'TW', WDID = @wdid, " +
                                                //    "EmpGender = @gender, ChangeBe = @change, EmpAdLoginID = @ad " +
                                                //    " WHERE EmpUniqNo = @no";
                                                cmdL.Parameters.Clear();
                                                cmdL.Parameters.Add(new SqlParameter("@locname", SqlDbType.NVarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@engname", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@startdate", SqlDbType.DateTime));
                                                cmdL.Parameters.Add(new SqlParameter("@quitdate", SqlDbType.DateTime));
                                                cmdL.Parameters.Add(new SqlParameter("@dept", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@job", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@email", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@home", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@mobile", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@boss", SqlDbType.Int));
                                                cmdL.Parameters.Add(new SqlParameter("@comp", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@next", SqlDbType.Int));
                                                cmdL.Parameters.Add(new SqlParameter("@no", SqlDbType.Int));
                                                //cmdL.Parameters.Add(new SqlParameter("@ext", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@wdid", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@gender", SqlDbType.Char));
                                                cmdL.Parameters.Add(new SqlParameter("@change", SqlDbType.Int));
                                                //cmdL.Parameters.Add(new SqlParameter("@ad", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@brand", SqlDbType.NVarChar));

                                                cmdL.Parameters["@locname"].Value = rdW.GetString(1).Trim();
                                                cmdL.Parameters["@engname"].Value = rdW.GetString(2).Trim();
                                                cmdL.Parameters["@startdate"].Value = DateTime.Parse(rdW.GetString(3));
                                                if (rdW.GetString(4).Trim() == "")
                                                    cmdL.Parameters["@quitdate"].Value = DBNull.Value;
                                                else
                                                    cmdL.Parameters["@quitdate"].Value = DateTime.Parse(rdW.GetString(4));
                                                cmdL.Parameters["@dept"].Value = rdW.GetString(5).Trim();
                                                cmdL.Parameters["@job"].Value = rdW.GetString(6).Trim();
                                                cmdL.Parameters["@email"].Value = rdW.GetString(7).Trim();
                                                cmdL.Parameters["@home"].Value = rdW.GetString(8).Trim();
                                                cmdL.Parameters["@mobile"].Value = rdW.GetString(9).Trim();
                                                cmdL.Parameters["@boss"].Value = GetBossUniqNo(rdW.GetString(12).Trim(), strCommon_L);
                                                cmdL.Parameters["@comp"].Value = rdW.GetString(11).Trim();
                                                cmdL.Parameters["@next"].Value = GetBossUniqNo(rdW.GetString(10).Trim(), strCommon_L);
                                                cmdL.Parameters["@no"].Value = iUniq;
                                                //cmdL.Parameters["@ext"].Value = rdW.GetString(13).Trim();
                                                cmdL.Parameters["@wdid"].Value = rdW.GetString(15).Trim();
                                                cmdL.Parameters["@gender"].Value = (rdW.GetString(14).Trim() == "1" ? "M" : "F");
                                                cmdL.Parameters["@change"].Value = -1;
                                                //if (rdW.GetString(16).Trim() == "") cmdL.Parameters["@ad"].Value = DBNull.Value;
                                                //else cmdL.Parameters["@ad"].Value = rdW.GetString(16).Trim();
                                                cmdL.Parameters["@brand"].Value = GetBrandName(rdW.GetString(16).Trim(), rdW.GetString(11).Trim(), rdW.GetString(5).Trim());

                                                cmdL.ExecuteNonQuery();

                                                if (rdW.GetString(4).Trim() != "")
                                                {
                                                    DANCom.DANClass.WriteToLog(PROG_ID, -1, "W", PROG_USER_ID, sComputerName, "員工 " + rdW.GetString(1).Trim() + "(" + rdW.GetString(2).Trim() + ") 於 " + rdW.GetString(4).Trim() + " 離職");
                                                }
                                                #endregion
                                            }
                                            else
                                            {
                                                rdL.Close();
                                            }
                                            #endregion  
                                        }
                                        else        //該員工編號資料不存在
                                        {
                                            rdL.Close();

                                            #region 以員工Workday ID查詢該員工資料是否已經存在
                                            //if (rdW.GetString(4).Trim() != "") continue;
                                            //                            0           1           2            3             4          5         6               7            8               9              10          11            12              13         14      15        16        17
                                            cmdL.CommandText = "SELECT EmpUniqNo, EmpLocName, EmpEngName, EmpStartDate, EmpQuitDate, DeptCode, JobTitleCode, EmpCompEmail, EmpHomePhone, EmpMobilePhone, BossEmpUniqNo, CompCode, NextBossEmpUniqNo, EmpExtNo, EmpGender, WDID, BrandCode, EmpNo FROM Employee WHERE RegionCode = 'TW' AND (EmpQuitDate IS NULL OR EmpQuitDate = '' OR EmpQuitDate >= GetDate()) AND WDID = @wdid";
                                            cmdL.Parameters.Clear();
                                            cmdL.Parameters.AddWithValue("@wdid", rdW.GetString(15));
                                            rdL = cmdL.ExecuteReader();
                                            if (rdL.HasRows)        //以員工中文姓名查詢員工資料已經存在
                                            {
                                                #region 以員工Workday ID查詢員工資料已經存在
                                                rdL.Read();
                                                //if (rdW.GetString(1).Trim() != rdL.GetString(1) ||
                                                //    rdW.GetString(2).Trim() != rdL.GetString(2) ||
                                                //    rdW.GetString(3).Trim() != rdL.GetDateTime(3).ToString("yyyy/MM/dd") ||
                                                //    rdW.GetString(4).Trim() != (rdL.IsDBNull(4) ? "" : rdL.GetDateTime(4).ToString("yyyy/MM/dd")) ||
                                                //    rdW.GetString(5).Trim() != (rdL.IsDBNull(5) ? "" : rdL.GetString(5)) ||
                                                //    rdW.GetString(6).Trim() != (rdL.IsDBNull(6) ? "" : rdL.GetString(6)) ||
                                                //    rdW.GetString(7).Trim() != (rdL.IsDBNull(7) ? "" : rdL.GetString(7)) ||
                                                //    rdW.GetString(8).Trim() != (rdL.IsDBNull(8) ? "" : rdL.GetString(8)) ||
                                                //    rdW.GetString(9).Trim() != (rdL.IsDBNull(9) ? "" : rdL.GetString(9)) ||
                                                //    rdW.GetString(10).Trim() != GetEmpNo((rdL.IsDBNull(10) ? 0 : rdL.GetInt32(10)), strCommon_L) ||
                                                //    (rdW.GetString(11).Trim() != (rdL.IsDBNull(11) ? "" : rdL.GetString(11)) && ((rdW.GetString(11).Trim() != "PO") || (rdL.GetString(11).Trim() != "AMP"))) ||
                                                //    rdW.GetString(12).Trim() != GetEmpNo((rdL.IsDBNull(12) ? 0 : rdL.GetInt32(12)), strCommon_L) ||
                                                //    rdW.GetString(13).Trim() != (rdL.IsDBNull(13) ? "" : rdL.GetString(13)) ||
                                                //    (rdW.GetString(14).Trim() == "1" ? "M" : "F") != (rdL.IsDBNull(14) ? "" : rdL.GetString(14)) ||
                                                //    rdW.GetString(15).Trim() != (rdL.IsDBNull(15) ? "" : rdL.GetString(15)) || 
                                                //    rdW.GetString(16).Trim() != (rdL.IsDBNull(16) ? "" : rdL.GetString(16)))
                                                if (rdW.GetString(1).Trim() != rdL.GetString(1) ||
                                                    rdW.GetString(2).Trim() != rdL.GetString(2) ||
                                                    rdW.GetString(3).Trim() != rdL.GetDateTime(3).ToString("yyyy/MM/dd") ||
                                                    rdW.GetString(4).Trim() != (rdL.IsDBNull(4) ? "" : rdL.GetDateTime(4).ToString("yyyy/MM/dd")) ||
                                                    rdW.GetString(5).Trim() != (rdL.IsDBNull(5) ? "" : rdL.GetString(5)) ||
                                                    rdW.GetString(6).Trim() != (rdL.IsDBNull(6) ? "" : rdL.GetString(6)) ||
                                                    rdW.GetString(7).Trim() != (rdL.IsDBNull(7) ? "" : rdL.GetString(7)) ||
                                                    rdW.GetString(8).Trim() != (rdL.IsDBNull(8) ? "" : rdL.GetString(8)) ||
                                                    rdW.GetString(9).Trim() != (rdL.IsDBNull(9) ? "" : rdL.GetString(9)) ||
                                                    rdW.GetString(10).Trim() != GetEmpNo((rdL.IsDBNull(12) ? 0 : rdL.GetInt32(12)), strCommon_L) ||
                                                    (rdW.GetString(11).Trim() != (rdL.IsDBNull(11) ? "" : rdL.GetString(11)) && ((rdW.GetString(11).Trim() != "PO") || (rdL.GetString(11).Trim() != "AMP"))) ||
                                                    rdW.GetString(12).Trim() != GetEmpNo((rdL.IsDBNull(10) ? 0 : rdL.GetInt32(10)), strCommon_L) ||
                                                    rdW.GetString(13).Trim() != (rdL.IsDBNull(13) ? "" : rdL.GetString(13)) ||
                                                    (rdW.GetString(14).Trim() == "1" ? "M" : "F") != (rdL.IsDBNull(14) ? "" : rdL.GetString(14)) ||
                                                    rdW.GetString(15).Trim() != (rdL.IsDBNull(15) ? "" : rdL.GetString(15)) ||
                                                    GetBrandName(rdW.GetString(16).Trim(), rdW.GetString(11).Trim(), rdW.GetString(5).Trim()) != (rdL.IsDBNull(16) ? "" : rdL.GetString(16)) ||
                                                    rdW.GetString(0).Trim() != rdL.GetString(17))
                                                {
                                                    int iUniq = rdL.GetInt32(0);
                                                    rdL.Close();

                                                    #region 員工資料有異動則更新
                                                    cmdL.CommandText = "UPDATE Employee SET EmpLocName = @locname, EmpEngName = @engname, " +
                                                        "EmpStartDate = @startdate, EmpQuitDate = @quitdate, DeptCode = @dept, CompCode = @comp, " +
                                                        "JobTitleCode = @job, EmpCompEmail = @email, EmpHomePhone = @home, EmpMobilePhone = @mobile, " +
                                                        "BossEmpUniqNo = @boss, NextBossEmpUniqNo = @next, RegionCode = 'TW', EmpNo = @eno, WDID = @wdid, " +
                                                        "EmpGender = @gender, BrandCode = @brand, ChangeBe = @change " +
                                                        " WHERE EmpUniqNo = @no";
                                                    //cmdL.CommandText = "UPDATE Employee SET EmpLocName = @locname, EmpEngName = @engname, " +
                                                    //    "EmpStartDate = @startdate, EmpQuitDate = @quitdate, DeptCode = @dept, CompCode = @comp, " +
                                                    //    "JobTitleCode = @job, EmpCompEmail = @email, EmpHomePhone = @home, EmpMobilePhone = @mobile, " +
                                                    //    "BossEmpUniqNo = @boss, NextBossEmpUniqNo = @next, RegionCode = 'TW', WDID = @wdid, " +
                                                    //    "EmpGender = @gender, ChangeBe = @change, EmpAdLoginID = @ad " +
                                                    //    " WHERE EmpUniqNo = @no";
                                                    cmdL.Parameters.Clear();
                                                    cmdL.Parameters.Add(new SqlParameter("@locname", SqlDbType.NVarChar));
                                                    cmdL.Parameters.Add(new SqlParameter("@engname", SqlDbType.VarChar));
                                                    cmdL.Parameters.Add(new SqlParameter("@startdate", SqlDbType.DateTime));
                                                    cmdL.Parameters.Add(new SqlParameter("@quitdate", SqlDbType.DateTime));
                                                    cmdL.Parameters.Add(new SqlParameter("@dept", SqlDbType.VarChar));
                                                    cmdL.Parameters.Add(new SqlParameter("@job", SqlDbType.VarChar));
                                                    cmdL.Parameters.Add(new SqlParameter("@email", SqlDbType.VarChar));
                                                    cmdL.Parameters.Add(new SqlParameter("@home", SqlDbType.VarChar));
                                                    cmdL.Parameters.Add(new SqlParameter("@mobile", SqlDbType.VarChar));
                                                    cmdL.Parameters.Add(new SqlParameter("@boss", SqlDbType.Int));
                                                    cmdL.Parameters.Add(new SqlParameter("@comp", SqlDbType.VarChar));
                                                    cmdL.Parameters.Add(new SqlParameter("@next", SqlDbType.Int));
                                                    cmdL.Parameters.Add(new SqlParameter("@eno", SqlDbType.VarChar));
                                                    cmdL.Parameters.Add(new SqlParameter("@no", SqlDbType.Int));
                                                    //cmdL.Parameters.Add(new SqlParameter("@ext", SqlDbType.VarChar));
                                                    cmdL.Parameters.Add(new SqlParameter("@wdid", SqlDbType.VarChar));
                                                    cmdL.Parameters.Add(new SqlParameter("@gender", SqlDbType.Char));
                                                    cmdL.Parameters.Add(new SqlParameter("@change", SqlDbType.Int));
                                                    //cmdL.Parameters.Add(new SqlParameter("@ad", SqlDbType.VarChar));
                                                    cmdL.Parameters.Add(new SqlParameter("@brand", SqlDbType.NVarChar));

                                                    cmdL.Parameters["@locname"].Value = rdW.GetString(1).Trim();
                                                    cmdL.Parameters["@engname"].Value = rdW.GetString(2).Trim();
                                                    cmdL.Parameters["@startdate"].Value = DateTime.Parse(rdW.GetString(3));
                                                    if (rdW.GetString(4).Trim() == "")
                                                        cmdL.Parameters["@quitdate"].Value = DBNull.Value;
                                                    else
                                                        cmdL.Parameters["@quitdate"].Value = DateTime.Parse(rdW.GetString(4));
                                                    cmdL.Parameters["@dept"].Value = rdW.GetString(5).Trim();
                                                    cmdL.Parameters["@job"].Value = rdW.GetString(6).Trim();
                                                    cmdL.Parameters["@email"].Value = rdW.GetString(7).Trim();
                                                    cmdL.Parameters["@home"].Value = rdW.GetString(8).Trim();
                                                    cmdL.Parameters["@mobile"].Value = rdW.GetString(9).Trim();
                                                    cmdL.Parameters["@boss"].Value = GetBossUniqNo(rdW.GetString(12).Trim(), strCommon_L);
                                                    cmdL.Parameters["@comp"].Value = rdW.GetString(11).Trim();
                                                    cmdL.Parameters["@next"].Value = GetBossUniqNo(rdW.GetString(10).Trim(), strCommon_L);
                                                    cmdL.Parameters["@eno"].Value = rdW.GetString(0);                                                   
                                                    cmdL.Parameters["@no"].Value = iUniq;
                                                    //cmdL.Parameters["@ext"].Value = rdW.GetString(13).Trim();
                                                    cmdL.Parameters["@wdid"].Value = rdW.GetString(15).Trim();
                                                    cmdL.Parameters["@gender"].Value = (rdW.GetString(14).Trim() == "1" ? "M" : "F");
                                                    cmdL.Parameters["@change"].Value = -1;
                                                    //if (rdW.GetString(16).Trim() == "") cmdL.Parameters["@ad"].Value = DBNull.Value;
                                                    //else cmdL.Parameters["@ad"].Value = rdW.GetString(16).Trim();
                                                    cmdL.Parameters["@brand"].Value = GetBrandName(rdW.GetString(16).Trim(), rdW.GetString(11).Trim(), rdW.GetString(5).Trim());

                                                    cmdL.ExecuteNonQuery();

                                                    if (rdW.GetString(4).Trim() != "")
                                                    {
                                                        DANCom.DANClass.WriteToLog(PROG_ID, -1, "W", PROG_USER_ID, sComputerName, "員工 " + rdW.GetString(1).Trim() + "(" + rdW.GetString(2).Trim() + ") 於 " + rdW.GetString(4).Trim() + " 離職");
                                                    }
                                                    #endregion
                                                }
                                                else
                                                {
                                                    rdL.Close();

                                                }
                                                #endregion
                                            }
                                            else
                                            {
                                                rdL.Close();

                                                #region 以員工編號、Workday ID都查不到該資料，所以新增一筆資料
                                                //cmdL.CommandText = "INSERT INTO Employee (EmpNo, EmpLocName, EmpEngName, EmpStartDate, EmpQuitDate, " +
                                                //    "DeptCode, JobTitleCode, EmpCompEmail, EmpHomePhone, EmpMobilePhone, BossEmpUniqNo, CompCode, " +
                                                //    "NextBossEmpUniqNo, EmpExtNo, RegionCode, EmpGender, OfficeCode, ChangeBe) VALUES (@no, @locname, @engname, @startdate, " +
                                                //    "@quitdate, @dept, @job, @email, @home, @mobile, @boss, @comp, @next, @ext, 'TW', @gender, @office, @change)";
                                                cmdL.CommandText = "INSERT INTO Employee (EmpNo, EmpLocName, EmpEngName, EmpStartDate, EmpQuitDate, " +
                                                    "DeptCode, JobTitleCode, EmpCompEmail, EmpHomePhone, EmpMobilePhone, BossEmpUniqNo, CompCode, " +
                                                    "NextBossEmpUniqNo, RegionCode, WDID, EmpGender, OfficeCode, ChangeBe, EmpType, BrandCode) VALUES (@no, @locname, @engname, @startdate, " +
                                                    "@quitdate, @dept, @job, @email, @home, @mobile, @boss, @comp, @next, 'TW', @wdid, @gender, @office, @change, '0', @brand)";
                                                cmdL.Parameters.Clear();
                                                cmdL.Parameters.Add(new SqlParameter("@no", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@locname", SqlDbType.NVarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@engname", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@startdate", SqlDbType.DateTime));
                                                cmdL.Parameters.Add(new SqlParameter("@quitdate", SqlDbType.DateTime));
                                                cmdL.Parameters.Add(new SqlParameter("@dept", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@job", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@email", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@home", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@mobile", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@boss", SqlDbType.Int));
                                                cmdL.Parameters.Add(new SqlParameter("@comp", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@next", SqlDbType.Int));
                                                //cmdL.Parameters.Add(new SqlParameter("@ext", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@wdid", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@gender", SqlDbType.Char));
                                                cmdL.Parameters.Add(new SqlParameter("@office", SqlDbType.VarChar));
                                                cmdL.Parameters.Add(new SqlParameter("@change", SqlDbType.Int));
                                                cmdL.Parameters.Add(new SqlParameter("@brand", SqlDbType.NVarChar));

                                                cmdL.Parameters["@no"].Value = rdW.GetString(0).Trim();
                                                cmdL.Parameters["@locname"].Value = rdW.GetString(1).Trim();
                                                cmdL.Parameters["@engname"].Value = rdW.GetString(2).Trim();
                                                cmdL.Parameters["@startdate"].Value = DateTime.Parse(rdW.GetString(3));
                                                if (rdW.GetString(4).Trim() == "")
                                                    cmdL.Parameters["@quitdate"].Value = DBNull.Value;
                                                else
                                                    cmdL.Parameters["@quitdate"].Value = DateTime.Parse(rdW.GetString(4));
                                                cmdL.Parameters["@dept"].Value = rdW.GetString(5).Trim();
                                                cmdL.Parameters["@job"].Value = rdW.GetString(6).Trim();
                                                cmdL.Parameters["@email"].Value = rdW.GetString(7).Trim() == "" ? GetEmail(rdW.GetString(15)) : rdW.GetString(7);
                                                cmdL.Parameters["@home"].Value = rdW.GetString(8).Trim();
                                                cmdL.Parameters["@mobile"].Value = rdW.GetString(9).Trim();                                                
                                                cmdL.Parameters["@boss"].Value = GetBossUniqNo(rdW.GetString(12).Trim(), strCommon_L);
                                                cmdL.Parameters["@comp"].Value = rdW.GetString(11).Trim();
                                                cmdL.Parameters["@next"].Value = GetBossUniqNo(rdW.GetString(10).Trim(), strCommon_L);
                                                //cmdL.Parameters["@ext"].Value = rdW.GetString(13).Trim();
                                                cmdL.Parameters["@wdid"].Value = rdW.GetString(15).Trim();
                                                cmdL.Parameters["@gender"].Value = (rdW.GetString(14).Trim() == "1" ? "M" : "F");
                                                cmdL.Parameters["@office"].Value = "TPE";
                                                cmdL.Parameters["@change"].Value = -1;
                                                cmdL.Parameters["@brand"].Value = GetBrandName(rdW.GetString(16).Trim(), rdW.GetString(11).Trim(), rdW.GetString(5).Trim());

                                                cmdL.ExecuteNonQuery();

                                                cmdL.CommandText = "SELECT @@IDENTITY";
                                                rdL = cmdL.ExecuteReader();
                                                rdL.Read();
                                                Decimal iNo = rdL.GetDecimal(0);
                                                rdL.Close();

                                                SetEmpPassword(iNo, rdW.GetString(0).Trim());

                                                if (rdW.GetString(4).Trim() != "")
                                                {
                                                    DANCom.DANClass.WriteToLog(PROG_ID, -1, "W", PROG_USER_ID, sComputerName, "員工 " + rdW.GetString(1).Trim() + "(" + rdW.GetString(2).Trim() + ") 於 " + rdW.GetString(4).Trim() + " 離職");
                                                }
                                                #endregion
                                            }
                                            #endregion
                                        }
                                    }
                                }
                            }
                        }
                        rdWL.Close();

                    }
                    //DANCom.DANClass.WriteToLog(PROG_ID, 3, "D", PROG_USER_ID, sComputerName, "員工基本資料轉檔結束");
                }
                catch (Exception ex)
                {
                    string sMsg = "員工基本資料轉檔錯誤" + sDBType + " CommonDB資料庫 - " + sEmpNo + " - " + sActive + " - " + ex.Message;
                    DANCom.DANClass.SendMail(RECIPIENT, "員工基本資料轉檔錯誤！", sMsg);
                    DANCom.DANClass.WriteToLog(PROG_ID, 3, "E", PROG_USER_ID, sComputerName, sMsg);
                    DANCom.DANClass.WriteToLog(PROG_ID, 3, "E", PROG_USER_ID, sComputerName, sStatment);
                }
            }
            #endregion

            #region 轉員工資料異動 Emptra
            if (bTran)
            {
                //DANCom.DANClass.WriteToLog(PROG_ID, 4, "D", PROG_USER_ID, sComputerName, "開始轉員工資料異動");
                string sStatment = "";
                string sEmpNo = "";

                try
                {
                    sDBType = "Wiselan";
                    using (SqlConnection cnWL = new SqlConnection(strWiselan))
                    {
                        cnWL.Open();
                        SqlDataAdapter ada = new SqlDataAdapter("SELECT * FROM Emptra WHERE crdate >= @crdate ORDER BY empno, crdate, seqno", cnWL);
                        //SqlDataAdapter ada = new SqlDataAdapter("SELECT * FROM Emptra WHERE crdate >= '2015/06/21' ORDER BY empno, crdate, seqno", cnWL);
                        ada.SelectCommand.Parameters.Clear();
                        ada.SelectCommand.Parameters.Add(new SqlParameter("@crdate", SqlDbType.VarChar));
                        ada.SelectCommand.Parameters["@crdate"].Value = DateTime.Today.AddDays(-7).ToString("yyyy/MM/dd");
                        DataSet ds = new DataSet();
                        ada.Fill(ds, "idList");

                        if (ds.Tables["idList"].Rows.Count > 0)
                        {
                            for (int i = 0; i < ds.Tables["idList"].Rows.Count; i++)
                            {
                                sEmpNo = ds.Tables["idList"].Rows[i]["empno"].ToString().Trim();
                                //if (sEmpNo == "T10311")
                                //    sEmpNo = ds.Tables["idList"].Rows[i]["empno"].ToString();
                                //if (sEmpNo == "9416" || sEmpNo == "9415")
                                //    sEmpNo = ds.Tables["idList"].Rows[i]["empno"].ToString().Trim();
                                SqlConnection cnL = new SqlConnection(strCommon_L);
                                SqlConnection cnL_Write = new SqlConnection(strCommon_L);
                                sDBType = "Local";
                                cnL.Open();
                                cnL_Write.Open();
                                SqlCommand cmdL = new SqlCommand();
                                cmdL.Connection = cnL;
                                SqlCommand cmdL_Write = new SqlCommand();
                                cmdL_Write.Connection = cnL_Write;
                                sDBType = "LocalWrite";
                                
                                
                                sDBType = "Local";
                                cmdL.CommandText = "SELECT EmpUniqNo, CompCode, DeptCode, JobTitleCode " +
                                    " FROM Employee WHERE RegionCode = 'TW' " +
                                    " AND (EmpQuitDate IS NULL OR EmpQuitDate = '' OR EmpQuitDate >= GetDate()) " +
                                    " AND EmpNo = @empno";
                                cmdL.Parameters.Clear();
                                cmdL.Parameters.Add(new SqlParameter("@empno", SqlDbType.VarChar));
                                cmdL.Parameters["@empno"].Value = ds.Tables["idList"].Rows[i]["empno"].ToString().Trim();
                                SqlDataReader rdL = cmdL.ExecuteReader();   //取得GCMIS遠端員工資料
                                if (rdL.HasRows)
                                {
                                    while (rdL.Read())
                                    {
                                        if (rdL.GetString(1).CompareTo(ds.Tables["idList"].Rows[i]["ncomno"].ToString().Trim()) != 0 ||
                                            rdL.GetString(2).CompareTo(ds.Tables["idList"].Rows[i]["ndepno"].ToString().Trim()) != 0 ||
                                            rdL.GetString(3).CompareTo(ds.Tables["idList"].Rows[i]["ntilno"].ToString().Trim()) != 0)
                                        {
                                            cmdL_Write.CommandText = "UPDATE Employee SET CompCode = @compcode, DeptCode = @deptcode, JobTitleCode = @titlecode, ChangeBe = @change WHERE EmpUniqNo = @uniqno";
                                            cmdL_Write.Parameters.Clear();
                                            cmdL_Write.Parameters.Add(new SqlParameter("@compcode", SqlDbType.VarChar));
                                            cmdL_Write.Parameters.Add(new SqlParameter("@deptcode", SqlDbType.VarChar));
                                            cmdL_Write.Parameters.Add(new SqlParameter("@titlecode", SqlDbType.VarChar));
                                            cmdL_Write.Parameters.Add(new SqlParameter("@uniqno", SqlDbType.Int));
                                            cmdL_Write.Parameters.Add(new SqlParameter("@change", SqlDbType.Int));

                                            cmdL_Write.Parameters["@compcode"].Value = ds.Tables["idList"].Rows[i]["ncomno"].ToString().Trim();
                                            cmdL_Write.Parameters["@deptcode"].Value = ds.Tables["idList"].Rows[i]["ndepno"].ToString().Trim();
                                            cmdL_Write.Parameters["@titlecode"].Value = ds.Tables["idList"].Rows[i]["ntilno"].ToString().Trim();
                                            cmdL_Write.Parameters["@uniqno"].Value = rdL.GetInt32(0);
                                            cmdL_Write.Parameters["@change"].Value = -2;
                                            cmdL_Write.ExecuteNonQuery();
                                        }
                                    }
                                }
                                rdL.Close();
                            }
                        }
                    }
                    //DANCom.DANClass.WriteToLog(PROG_ID, 4, "D", PROG_USER_ID, sComputerName, "員工資料異動轉檔結束");
                }
                catch (Exception ex)
                {
                    string sMsg = "員工資料異動轉檔錯誤" + sDBType + " CommonDB資料庫 - " + sEmpNo + " - " + sActive + " - " + ex.Message;
                    DANCom.DANClass.SendMail(RECIPIENT, "員工資料異動轉檔錯誤！", sMsg);
                    DANCom.DANClass.WriteToLog(PROG_ID, 4, "E", PROG_USER_ID, sComputerName, sMsg);
                    DANCom.DANClass.WriteToLog(PROG_ID, 4, "E", PROG_USER_ID, sComputerName, sStatment);
                }
            }
            #endregion

            #region 轉員工主管資料
            if (bBoss)
            {
                //DANCom.DANClass.WriteToLog(PROG_ID, 4, "D", PROG_USER_ID, sComputerName, "開始轉員工主管資料");
                string dtToday = DateTime.Today.ToString("yyyy/MM/dd");

                sDBType = "Wiselan";
                //更新員工的二階主管
                using (SqlConnection cnWL = new SqlConnection(strWiselan))
                {
                    try
                    {
                        cnWL.Open();
                        //directno - 一階主管  topman - 二階主管
                        SqlCommand cmdWL = new SqlCommand("SELECT empno, directno, topman, comno FROM Empmsf WHERE odat = '' OR odat >= '" + dtToday + "'", cnWL);
                        SqlDataReader rdWL = cmdWL.ExecuteReader();
                        if (rdWL.HasRows)
                        {
                            #region 使用的變數
                            Int32 iUniqNo = 0;
                            Int32 iFirst = 0;       //一階主管總號（從Master Data取得）
                            string sFirstEngName;
                            string sFirstLocName;
                            Int32 iSecond = 0;      //二階主管總號（從Master Data取得）
                            string sSecondEngName;
                            string sSecondLocName;
                            string sDirect;         //一階主管員編
                            string sTopMan;         //二階主管員編
                            string sEmpEngName;
                            string sEmpLocName;
                            Int32 iBoss = 0;        //從智將取得的一階主管總號
                            Int32 iSecondBoss = 0;  //從智將取得的二階主管總號
                            #endregion

                            SqlConnection cnL = new SqlConnection(strCommon_L);
                            SqlConnection cnMIS = new SqlConnection(strGCMIS);
                            sDBType = "Local";
                            cnL.Open();
                            cnMIS.Open();
                            SqlCommand cmdL = new SqlCommand();
                            SqlCommand cmdMIS = new SqlCommand();
                            cmdL.Connection = cnL;
                            cmdMIS.Connection = cnMIS;

                            SqlDataReader rdL;
                            //SqlDataReader rdMIS;

                            while (rdWL.Read())
                            {
                                //if (rdWL.GetString(0).Trim() == "9852")
                                //    sDBType = "Local";
                                if (rdWL.GetString(3).Trim() == "DX" || rdWL.GetString(3).Trim() == "DT") return; //如果是台灣電通或者是貝立德則不予處理
                                #region 先存主管的員編
                                sDirect = rdWL.GetString(1).Trim();
                                sTopMan = rdWL.GetString(2).Trim();
                                #endregion

                                #region 取得主管的總號
                                iBoss = GetBossUniqNo(sDirect, strCommon_L);
                                iSecondBoss = GetBossUniqNo(sTopMan, strCommon_L);
                                #endregion

                                sActive = "Read1";
                                cmdL.CommandText = "SELECT EmpUniqNo, EmpNo, BossEmpUniqNo, NextBossEmpUniqNo, EmpEngName, EmpLocName FROM Employee WHERE RegionCode = 'TW' AND EmpNo = @no AND (EmpQuitDate IS NULL OR EmpQuitDate = '' OR EmpQuitDate >= GetDate())";
                                cmdL.Parameters.Clear();
                                cmdL.Parameters.Add(new SqlParameter("@no", SqlDbType.VarChar));
                                cmdL.Parameters["@no"].Value = rdWL.GetString(0).Trim();
                                rdL = cmdL.ExecuteReader();
                                if (rdL.HasRows)
                                {
                                    rdL.Read();
                                    iUniqNo = rdL.GetInt32(0);
                                    iFirst = (rdL.IsDBNull(2) ? 0 : rdL.GetInt32(2));
                                    iSecond = (rdL.IsDBNull(3) ? 0 : rdL.GetInt32(3));
                                    sEmpEngName = rdL.GetString(4);
                                    sEmpLocName = rdL.GetString(5);
                                    rdL.Close();
                                    //sActive = "Read2";
                                    //cmdL.CommandText = "SELECT EmpUniqNo, EmpEngName, EmpLocName FROM Employee WHERE RegionCode = 'TW' AND EmpNo = @direct AND (EmpQuitDate IS NULL OR EmpQuitDate = '' OR EmpQuitDate >= GetDate())";
                                    ////cmdL.CommandText = "SELECT EmpUniqNo, EmpEngName, EmpLocName FROM Employee WHERE RegionCode = 'TW' AND EmpNo = '" + sDirect + "'";
                                    //cmdL.Parameters.Clear();
                                    //cmdL.Parameters.Add(new SqlParameter("@direct", SqlDbType.VarChar));
                                    //cmdL.Parameters["@direct"].Value = sDirect;
                                    //rdL = cmdL.ExecuteReader();
                                    //if (rdL.HasRows)
                                    //{
                                    //    rdL.Read();
                                    //    iFirst = rdL.GetInt32(0);
                                    //    sFirstEngName = rdL.GetString(1);
                                    //    sFirstLocName = rdL.GetString(2);
                                    //}
                                    //else
                                    //{
                                    //    iFirst = 0;
                                    //    sFirstEngName = "";
                                    //    sFirstLocName = "";
                                    //}
                                    //rdL.Close();
                                    
                                    //sActive = "Read3";
                                    //cmdL.CommandText = "SELECT EmpUniqNo, EmpEngName, EmpLocName FROM Employee WHERE RegionCode = 'TW' AND EmpNo = @top AND (EmpQuitDate IS NULL OR EmpQuitDate = '' OR EmpQuitDate >= GetDate())";
                                    ////cmdL.CommandText = "SELECT EmpUniqNo, EmpEngName, EmpLocName FROM Employee WHERE RegionCode = 'TW' AND EmpNo = '" + sTopMan + "'";
                                    //cmdL.Parameters.Clear();
                                    //cmdL.Parameters.Add(new SqlParameter("@top", SqlDbType.VarChar));
                                    //cmdL.Parameters["@top"].Value = sTopMan;
                                    //rdL = cmdL.ExecuteReader();
                                    //if (rdL.HasRows)
                                    //{
                                    //    rdL.Read();
                                    //    iSecond = rdL.GetInt32(0);
                                    //    sSecondEngName = rdL.GetString(1);
                                    //    sSecondLocName = rdL.GetString(2);
                                    //}
                                    //else
                                    //{
                                    //    iSecond = 0;
                                    //    sSecondEngName = "";
                                    //    sSecondLocName = "";
                                    //}
                                    //rdL.Close();

                                    if (iBoss != iFirst || iSecondBoss != iSecond)
                                    {
                                        sActive = "Update1";
                                        cmdL.CommandText = "UPDATE Employee SET BossEmpUniqNo = @first, NextBossEmpUniqNo = @second" +
                                            " WHERE EmpUniqNo = @no";
                                        cmdL.Parameters.Clear();
                                        cmdL.Parameters.Add(new SqlParameter("@first", SqlDbType.Int));
                                        cmdL.Parameters.Add(new SqlParameter("@second", SqlDbType.Int));
                                        cmdL.Parameters.Add(new SqlParameter("@no", SqlDbType.Int));
                                        cmdL.Parameters["@first"].Value = iBoss;
                                        cmdL.Parameters["@second"].Value = iSecondBoss;
                                        cmdL.Parameters["@no"].Value = iUniqNo;
                                        cmdL.ExecuteNonQuery();

                                        try
                                        {
                                            sActive = "Update2";
                                            cmdMIS.CommandText = "UPDATE TicketDetail SET BossUniqNo = @bossno WHERE EmpUniqNo = @no AND CurrentStep = '3'";
                                            cmdMIS.Parameters.Clear();
                                            cmdMIS.Parameters.Add(new SqlParameter("@bossno", SqlDbType.Int));
                                            cmdMIS.Parameters.Add(new SqlParameter("@no", SqlDbType.Int));
                                            cmdMIS.Parameters["@bossno"].Value = (iBoss == 0 ? (iSecondBoss == 0 ? "null" : iSecondBoss.ToString()) : iBoss.ToString());
                                            cmdMIS.Parameters["@no"].Value = iUniqNo;
                                            cmdMIS.ExecuteNonQuery();
                                        }
                                        catch (Exception ex)
                                        {
                                            DANCom.DANClass.WriteToLog(PROG_ID, 4, "E", PROG_USER_ID, sComputerName, "iFirst = " + iBoss + "\niSecond = " + iSecondBoss + "\niBoss = " + iBoss + "\niSecondBoss = " + iSecondBoss + " - " + ex.Message);
                                        }
                                    }
                                }
                                else
                                    rdL.Close();
                            }

                        }
                        rdWL.Close();
                        //DANCom.DANClass.WriteToLog(PROG_ID, 4, "D", PROG_USER_ID, sComputerName, "員工主管資料轉檔結束");
                    }
                    catch (Exception ex)
                    {
                        DANCom.DANClass.SendMail(RECIPIENT, "員工主管資料轉檔錯誤！", @"員工主管資料轉檔錯誤 - " + sDBType + " - " + sActive + " - " + ex.Message);

                        DANCom.DANClass.WriteToLog(PROG_ID, 4, "E", PROG_USER_ID, sComputerName, sDBType + " - " + sActive + " - " + ex.Message);
                    }
                }
            }
            #endregion

            ////////////////////////////////////////////////
            
            #region 轉員工分機資料
            if (bExt)
            {
                DataTable dt = new DataTable();

                sDBType = "CommonDB";
                using (SqlConnection cnCMN = new SqlConnection(strCommon_L))
                {
                    try
                    {
                        cnCMN.Open();
                        #region 取得 CommonDB 目前在職人員的資料
                        SqlDataAdapter ada = new SqlDataAdapter("SELECT WDID, EmpUniqNo FROM Employee WHERE RegionCode = 'TW' AND (EmpQuitDate IS NULL OR EmpQuitDate = '' OR EmpQuitDate >= GetDate()) AND (EmpExtNo IS NULL OR EmpExtNo = '' OR LeaseLine IS NULL OR LeaseLine = '')", cnCMN);
                        ada.Fill(dt);
                        SqlCommand cmdL = new SqlCommand();
                        cmdL.Connection = cnCMN;
                        if (dt.Rows.Count > 0)
                        {
                            using (SqlConnection cnH = new SqlConnection(strHai))  //連結JML資料庫
                            {
                                cnH.Open();
                                SqlCommand cmdH = new SqlCommand();
                                cmdH.Connection = cnH;
                                SqlDataReader rdH;
                                for (int i = 0; i < dt.Rows.Count; i++)
                                {
                                    cmdH.CommandText = "SELECT DIDNo FROM DIDData WHERE UseWDID = @wdid";
                                    cmdH.Parameters.Clear();
                                    cmdH.Parameters.AddWithValue("@wdid", dt.Rows[i]["WDID"].ToString());
                                    rdH = cmdH.ExecuteReader();
                                    if (rdH.HasRows)
                                    {
                                        rdH.Read();
                                        cmdL.CommandText = "UPDATE Employee SET EmpExtNo = @ext, LeaseLine = @line WHERE EmpUniqNo = @uniq";
                                        cmdL.Parameters.Clear();
                                        cmdL.Parameters.AddWithValue("@ext", rdH.GetString(0).Substring(4));
                                        cmdL.Parameters.AddWithValue("@line", rdH.GetString(0));
                                        cmdL.Parameters.AddWithValue("@uniq", Int32.Parse(dt.Rows[i]["EmpUniqNo"].ToString()));
                                        cmdL.ExecuteNonQuery();
                                    }
                                    rdH.Close();
                                }
                            }
                        }
                        #endregion
                    }
                    catch (Exception ex)
                    {
                        DANCom.DANClass.SendMail(RECIPIENT, "員工分機資料轉檔錯誤！", @"員工分機資料轉檔錯誤" + sDBType + " - " + ex.Message);
                        DANCom.DANClass.WriteToLog(PROG_ID, 5, "E", PROG_USER_ID, sComputerName, sDBType + " - " + ex.Message);
                    }
                }
            }
            #endregion

            #region 轉員工耳機資料
            if (bEar)
            {
                DataTable dt = new DataTable();

                sDBType = "CommonDB";
                using (SqlConnection cnCMN = new SqlConnection(strCommon_L))
                {
                    try
                    {
                        cnCMN.Open();
                        #region 取得 CommonDB 目前在職人員的資料
                        SqlDataAdapter ada = new SqlDataAdapter("SELECT WDID, EmpUniqNo FROM Employee WHERE RegionCode = 'TW' AND (EmpQuitDate IS NULL OR EmpQuitDate = '' OR EmpQuitDate >= GetDate()) AND (AssestNo IS NULL OR AssestNo = '')", cnCMN);
                        ada.Fill(dt);
                        SqlCommand cmdL = new SqlCommand();
                        cmdL.Connection = cnCMN;
                        if (dt.Rows.Count > 0)
                        {
                            using (SqlConnection cnH = new SqlConnection(strHai))  //連結JML資料庫
                            {
                                cnH.Open();
                                SqlCommand cmdH = new SqlCommand();
                                cmdH.Connection = cnH;
                                SqlDataReader rdH;
                                for (int i = 0; i < dt.Rows.Count; i++)
                                {
                                    cmdH.CommandText = "SELECT AssestNo, ThirdCode FROM EarPhoneData WHERE WDID = @wdid";
                                    cmdH.Parameters.Clear();
                                    cmdH.Parameters.AddWithValue("@wdid", dt.Rows[i]["WDID"].ToString());
                                    rdH = cmdH.ExecuteReader();
                                    if (rdH.HasRows)
                                    {
                                        rdH.Read();
                                        cmdL.CommandText = "UPDATE Employee SET AssestNo = @no, ThirdCode = @code WHERE EmpUniqNo = @uniq";
                                        cmdL.Parameters.Clear();
                                        cmdL.Parameters.AddWithValue("@no", rdH.IsDBNull(0) ? "" : rdH.GetString(0));
                                        cmdL.Parameters.AddWithValue("@code", rdH.IsDBNull(1) ? "" : rdH.GetString(1));
                                        cmdL.Parameters.AddWithValue("@uniq", Int32.Parse(dt.Rows[i]["EmpUniqNo"].ToString()));
                                        cmdL.ExecuteNonQuery();
                                    }
                                    rdH.Close();
                                }
                            }
                        }
                        #endregion
                    }
                    catch (Exception ex)
                    {
                        DANCom.DANClass.SendMail(RECIPIENT, "員工耳機資料轉檔錯誤！", @"員工耳機資料轉檔錯誤" + sDBType + " - " + ex.Message);
                        DANCom.DANClass.WriteToLog(PROG_ID, 5, "E", PROG_USER_ID, sComputerName, sDBType + " - " + ex.Message);
                    }
                }
                //sDBType = "Wiselan";
                //using (SqlConnection cnWL = new SqlConnection(strWiselan))
                //{
                //    try
                //    {
                //        cnWL.Open();

                //        SqlCommand cmdWL = new SqlCommand("SELECT empno, tel_no FROM Empmsf WHERE (odat = '' OR odat >= '" + DateTime.Today.ToString("yyyy/MM/dd") + "')", cnWL);
                //        SqlDataReader rdWL = cmdWL.ExecuteReader();
                //        if (rdWL.HasRows)
                //        {
                //            SqlConnection cnCmn_L = new SqlConnection(strCommon_L);
                //            sDBType = "MIS";
                //            SqlConnection cnMIS = new SqlConnection(strGCMIS);
                //            sDBType = "Local";
                //            cnCmn_L.Open();
                //            cnMIS.Open();
                //            SqlCommand cmdCmn_L = new SqlCommand();
                //            SqlCommand cmdMIS = new SqlCommand();
                //            cmdCmn_L.Connection = cnCmn_L;
                //            cmdMIS.Connection = cnMIS;
                //            SqlDataReader rdCmn_L;
                //            SqlDataReader rdMIS;

                //            while (rdWL.Read())
                //            {
                //                sDBType = "Local";
                //                cmdCmn_L.CommandText = "SELECT ISNULL(EmpExtNo, '') FROM Employee WHERE RegionCode = 'TW' AND EmpNo = '" + rdWL.GetString(0).Trim() + "' AND (EmpQuitDate IS NULL OR EmpQuitDate = '' OR EmpQuitDate >= GetDate())";
                //                rdCmn_L = cmdCmn_L.ExecuteReader();
                //                if (rdCmn_L.HasRows)
                //                {
                //                    rdCmn_L.Read();
                //                    sExt = rdCmn_L.GetString(0);
                //                    rdCmn_L.Close();
                //                    if (sExt.CompareTo(rdWL.GetString(1).Trim()) != 0)
                //                    {
                //                        cmdCmn_L.CommandText = "UPDATE Employee SET EmpExtNo = '" + rdWL.GetString(1).Trim() + "' WHERE RegionCode = 'TW' AND EmpNo = '" + rdWL.GetString(0).Trim() + "' AND (EmpQuitDate IS NULL OR EmpQuitDate >= GetDate())";
                //                        cmdCmn_L.ExecuteNonQuery();
                //                    }
                //                }
                //                else
                //                {
                //                    rdCmn_L.Close();
                //                }
                //            }
                //        }
                //        rdWL.Close();
                //    }
                //    catch (Exception ex)
                //    {
                //        DANCom.DANClass.SendMail(RECIPIENT, "員工分機資料轉檔錯誤！", @"員工分機資料轉檔錯誤" + sDBType + " - " + ex.Message);
                //        DANCom.DANClass.WriteToLog(PROG_ID, 5, "E", PROG_USER_ID, sComputerName, sDBType + " - " + ex.Message);
                //    }
                //}
            }
            #endregion

            #region 轉員工出勤資料
            if (bAttendance)
            {
                //DANCom.DANClass.WriteToLog(PROG_ID, 6, "D", PROG_USER_ID, sComputerName, "開始轉員工出勤資料");
                #region 正常情況下之轉檔－直接轉出勤資料
                sDBType = "Wiselan";
                using (SqlConnection cnWL = new SqlConnection(strWiselan))
                {
                    string sEmpNo = "";

                    try
                    {
                        cnWL.Open();

                        SqlCommand cmdWL = new SqlCommand("SELECT wdat, empno, time1, time2, wtime1, wtime2, daytyp FROM History WHERE wdat >= '" + DateTime.Today.AddDays(-7).ToString("yyyy/MM/dd") + "'", cnWL);
                        //SqlCommand cmdWL = new SqlCommand("SELECT wdat, empno, time1, time2, wtime1, wtime2, daytyp FROM History WHERE wdat >= '2018/04/25'", cnWL);
                        SqlDataReader rdWL = cmdWL.ExecuteReader();
                        SqlConnection cnOff = new SqlConnection(strWiselan);
                        cnOff.Open();
                        SqlCommand cmdOff = new SqlCommand();
                        cmdOff.Connection = cnOff;
                        SqlDataReader rdOff;

                        if (rdWL.HasRows)
                        {
                            int i = 0;

                            try
                            {
                                //string abc;
                                const int NOONSPAN = 60;
                                string sInTime;
                                string sOutTime;
                                int iOffSpan = 0;
                                int iSpan = 0;
                                int iOver = 0;
                                string isHoliday;

                                SqlConnection cnMIS = new SqlConnection(strGCMIS);
                                cnMIS.Open();

                                SqlCommand cmdMIS = new SqlCommand();
                                cmdMIS.Connection = cnMIS;
                                SqlDataReader rdMIS;

                                string sCompCode;
                                string sEmpEngName;
                                string sEmpLocName;
                                Int32 iEmpUniqNo;
                                string sBossEngName;
                                string sBossLocName;
                                string sDeptCode;
                                string sDeptName;
                                string sOfficeCode;
                                string sOnDutyTime;
                                string sOffDutyTime;

                                while (rdWL.Read())
                                {
                                    sDBType = "MIS";
                                    //2015/11/02 Robert 新增
                                    //Isobar 吳宛蓉(Maggie Wu) 2015/10/31離職，但繼續留下來支援，故其出勤資料予以忽略
                                    //if (rdWL.GetString(1).Trim() == "8686")
                                    //    continue;
                                    //if (rdWL.GetString(1).Trim() == "9739")
                                    //    sDBType = "MIS";

                                    //2016/05/05 Robert 新增
                                    //因為有些兼職員工轉正職後，卡機資料不會更新，因此會對應不到新的員工編號，如此更新後就不會再顯示錯誤訊息
                                    if (!CheckEmployee(rdWL.GetString(1), rdWL.GetString(0))) continue;

                                    sEmpNo = rdWL.GetString(1).Trim();
                                    //if (sEmpNo == "9880")
                                    //    sDBType = "MIS";
                                    
                                    i++;
                                    cmdOff.CommandText = "SELECT offtime FROM Offtra WHERE empno = '" + rdWL.GetString(1).Trim() + "' AND wdat = '" + rdWL.GetString(0).Trim() + "'";
                                    rdOff = cmdOff.ExecuteReader();
                                    if (rdOff.HasRows)
                                    {
                                        rdOff.Read();
                                        iOffSpan = (int)rdOff.GetDouble(0) * 60;
                                    }
                                    else
                                        iOffSpan = 0;
                                    rdOff.Close();

                                    cmdOff.CommandText = "SELECT * FROM Holday WHERE wdat = '" + rdWL.GetString(0).Trim() + "'";
                                    rdOff = cmdOff.ExecuteReader();
                                    if (rdOff.HasRows)
                                        isHoliday = "Y";
                                    else
                                        isHoliday = "N";
                                    rdOff.Close();

                                    if (rdWL.GetString(4).Trim() != "")
                                        sInTime = rdWL.GetString(4).Trim().Substring(0, 2) + ":" + rdWL.GetString(4).Trim().Substring(2) + ":00";
                                    else
                                        sInTime = null;

                                    if (rdWL.GetString(5).Trim() != "")
                                    {
                                        if (rdWL.GetString(6).Trim() == "Y")
                                        {
                                            sOutTime = DateTime.Parse(rdWL.GetString(0).Trim()).AddDays(1).ToString("yyyy/MM/dd") + " " + rdWL.GetString(5).Trim().Substring(0, 2) + ":" + rdWL.GetString(5).Trim().Substring(2) + ":00";
                                        }
                                        else
                                            sOutTime = rdWL.GetString(0).Trim() + " " + rdWL.GetString(5).Trim().Substring(0, 2) + ":" + rdWL.GetString(5).Trim().Substring(2) + ":00";

                                        if (rdWL.GetString(4).Trim() != "")
                                            if (isHoliday == "N")
                                                iSpan = (int)(DateTime.Parse(sOutTime) - DateTime.Parse(rdWL.GetString(0).Trim() + " " + sInTime)).TotalMinutes + iOffSpan - NOONSPAN;
                                            else
                                                iSpan = (int)(DateTime.Parse(sOutTime) - DateTime.Parse(rdWL.GetString(0).Trim() + " " + sInTime)).TotalMinutes + iOffSpan;
                                        else
                                            iSpan = 0;
                                    }
                                    else
                                    {
                                        sOutTime = null;
                                        iSpan = 0;
                                    }

                                    if (iSpan > 480)
                                        iOver = iSpan - 480;
                                    else
                                        iOver = 0;

                                    string sAllowSupper = null;
                                    string sAllowTaxi = null;

                                    if (isHoliday == "Y")   //假日
                                    {
                                        if (iSpan >= 480)
                                            sAllowSupper = "YES(2)";
                                        else
                                        {
                                            if (iSpan >= 240)
                                                sAllowSupper = "YES";
                                        }

                                        if (iSpan >= 360)
                                        {
                                            sAllowTaxi = "YES";
                                        }
                                    }
                                    else                    //上班日
                                    {
                                        if (iSpan >= 600 && (sOutTime.CompareTo(rdWL.GetString(0).Trim() + " 20:00:00") >= 0))
                                            sAllowSupper = "YES";
                                        else
                                            sAllowSupper = null;

                                        if (iSpan >= 720 && (sOutTime.CompareTo(rdWL.GetString(0).Trim() + " 22:00:00") >= 0))
                                            sAllowTaxi = "YES";
                                        else
                                            sAllowTaxi = null;
                                    }

                                    getEmpData(rdWL.GetString(1).Trim(), rdWL.GetString(0).Trim(), out sEmpEngName, out sEmpLocName, out sCompCode, out iEmpUniqNo, out sBossEngName, out sBossLocName, out sDeptCode, out sDeptName, out sOfficeCode);
                                    GetDutyTime(iEmpUniqNo, out sOnDutyTime, out sOffDutyTime);
                                    //cmdMIS.CommandText = "SELECT ISNULL(AllowTaxi, ''), ISNULL(AllowSupper, ''), ISNULL(ConfirmInTime, ''), LastCardTime, CheckOutTime, ConfirmOutTime, ISNULL(TotalWorkTime, 0), ISNULL(OverTime, 0) FROM EmpAttendance WHERE RegionCode = 'TW' AND WorkDate = '" + rdWL.GetString(0).Trim() + "' AND EmpNo = '" + rdWL.GetString(1).Trim() + "' AND CompCode = '" + sCompCode + "' AND OfficeCode = '" + sOfficeCode + "' AND EmpUniqNo = " + iEmpUniqNo.ToString();
                                    cmdMIS.CommandText = "SELECT ISNULL(AllowTaxi, ''), ISNULL(AllowSupper, ''), ISNULL(ConfirmInTime, ''), LastCardTime, CheckOutTime, ConfirmOutTime, ISNULL(TotalWorkTime, 0), ISNULL(OverTime, 0) FROM EmpAttendance WHERE RegionCode = 'TW' AND WorkDate = @date AND CompCode = @comp AND OfficeCode = @office AND EmpUniqNo = @uniq";
                                    cmdMIS.Parameters.Clear();
                                    cmdMIS.Parameters.AddWithValue("@date", rdWL.GetString(0).Trim());
                                    //cmdMIS.Parameters.AddWithValue("@no", rdWL.GetString(1).Trim());
                                    cmdMIS.Parameters.AddWithValue("@comp", sCompCode);
                                    cmdMIS.Parameters.AddWithValue("@office", sOfficeCode);
                                    cmdMIS.Parameters.AddWithValue("@uniq", iEmpUniqNo);
                                    rdMIS = cmdMIS.ExecuteReader();
                                    if (rdMIS.HasRows)
                                    {
                                        rdMIS.Read();
                                        
                                        sAllowTaxi = (rdMIS.GetString(0).Trim() == "" ? sAllowTaxi : rdMIS.GetString(0).Trim());
                                        sAllowSupper = (rdMIS.GetString(1).Trim() == "" ? sAllowSupper : rdMIS.GetString(1).Trim());

                                        if (rdMIS.GetString(0).Trim().CompareTo(sAllowTaxi == null ? "" : sAllowTaxi) != 0 ||
                                            rdMIS.GetString(1).Trim().CompareTo(sAllowSupper == null ? "" : sAllowSupper) != 0 ||
                                            rdMIS.GetString(2).Trim().CompareTo(sInTime) != 0 ||
                                            (rdMIS.IsDBNull(3) ? "9999/12/31 23:59:59" : rdMIS.GetDateTime(3).ToString("yyyy/MM/dd HH:mm:ss")) != (sOutTime == null ? "9999/12/31 23:59:59" : sOutTime) ||
                                            (rdMIS.IsDBNull(5) ? "9999/12/31 23:59:59" : rdMIS.GetDateTime(5).ToString("yyyy/MM/dd HH:mm:ss")) != (sOutTime == null ? "9999/12/31 23:59:59" : sOutTime) ||
                                            rdMIS.GetInt32(6) != iSpan || rdMIS.GetInt32(7) != iOver)
                                        {
                                            rdMIS.Close();

                                            cmdMIS.CommandText = "UPDATE EmpAttendance SET RegionCode = 'TW'," +
                                                " CompCode = '" + sCompCode + "', OfficeCode = 'TPE', " +
                                                "EmpUniqNo = " + iEmpUniqNo.ToString() + 
                                                ", WorkDate = '" + rdWL.GetString(0).Trim() + 
                                                "', EmpNo = '" + rdWL.GetString(1).Trim() + 
                                                "', EmpEngName = '" + sEmpEngName + 
                                                "', EmpLocName = N'" + sEmpLocName + 
                                                "', FirstCardTime = " + (rdWL.GetString(4).Trim() == "" ? "null" : "'" + sInTime + "'") + 
                                                ", InTime = " + (rdWL.GetString(4).Trim() == "" ? "null" : "'" + sInTime + "'") + 
                                                ", ConfirmInTime = " + (rdWL.GetString(4).Trim() == "" ? "null" : "'" + sInTime + "'") + 
                                                ", WorkBegTime = '" + sOnDutyTime + 
                                                "', LastCardTime = " + (sOutTime == null ? "null" : "'" + sOutTime + "'") +
                                                ", ConfirmOutTime = " + (sOutTime == null ? "null" : "'" + sOutTime + "'") +
                                                ", WorkOffTime = '" + rdWL.GetString(0).Trim() + " " + sOffDutyTime + 
                                                "', IsHoliday = " + (isHoliday == "Y" ? "'Y'" : "null") +
                                                ", BossEngName = " + (sBossEngName == "" ? "null" : "'" + sBossEngName + "'") + 
                                                ", BossLocName = " + (sBossLocName == "" ? "null" : "'" + sBossLocName + "'") + 
                                                ", TotalWorkTime = " + (sOutTime == null ? "null" : iSpan.ToString()) + 
                                                ", LeaveSpanTime = " + (sOutTime == null ? "null" : iOffSpan.ToString()) +
                                                ", OverTime = " + (iOver == 0 ? "null" : iOver.ToString()) + 
                                                ", AllowSupper = " + (sAllowSupper == null ? "null" : "'" + sAllowSupper.Trim() + "'") + 
                                                ", AllowTaxi = " + (sAllowTaxi == null ? "null" : "'" + sAllowTaxi.Trim() + "'") + 
                                                ", DeptCode = " + (sDeptCode == "" ? "null" : "'" + sDeptCode + "'") +
                                                ", DeptName = " + (sDeptName == "" ? "null" : "'" + sDeptName + "'") + 
                                                ", ChangeBy = -2, ChangeDate = getdate() WHERE RegionCode = 'TW' AND WorkDate = '" +
                                                rdWL.GetString(0).Trim() + "' AND CompCode = '" + sCompCode + "' AND OfficeCode = 'TPE' AND EmpUniqNo = " + iEmpUniqNo.ToString();
                                            cmdMIS.ExecuteNonQuery();
                                        }
                                        else 
                                        {
                                            rdMIS.Close();
                                        }
                                    }
                                    else
                                    {
                                        rdMIS.Close();

                                        getEmpData(rdWL.GetString(1).Trim(), rdWL.GetString(0).Trim(), out sEmpEngName, out sEmpLocName, out sCompCode, out iEmpUniqNo, out sBossEngName, out sBossLocName, out sDeptCode, out sDeptName, out sOfficeCode);

                                        cmdMIS.CommandText = "INSERT INTO EmpAttendance (RegionCode, CompCode, " +
                                            "OfficeCode, EmpUniqNo, WorkDate, EmpNo, EmpEngName, EmpLocName, " +
                                            "FirstCardTime, InTime, ConfirmInTime, WorkBegTime, LastCardTime, ConfirmOutTime, WorkOffTime, IsHoliday, " + 
                                            "BossEngName, BossLocName, TotalWorkTime, LeaveSpanTime, " +
                                            "OverTime, AllowSupper, AllowTaxi, DeptCode, DeptName, InputDate, ChangeBy, ChangeDate) " +
                                            " VALUES ('TW', '" + 
                                            sCompCode + "', 'TPE', " + iEmpUniqNo.ToString() + ", '" + 
                                            rdWL.GetString(0).Trim() + "', '" + rdWL.GetString(1).Trim() + "', '" + 
                                            sEmpEngName + "', N'" + sEmpLocName + "', " + 
                                            (rdWL.GetString(4).Trim() == "" ? "null" : "'" + sInTime + "'") + ", " +
                                            (rdWL.GetString(4).Trim() == "" ? "null" : "'" + sInTime + "'") + ", " + 
                                            (rdWL.GetString(4).Trim() == "" ? "null" : "'" + sInTime + "'") + ", '" +
                                            sOnDutyTime + "', " + (sOutTime == null ? "null" : "'" + sOutTime + "'") + ", " +
                                            (sOutTime == null ? "null" : "'" + sOutTime + "'") + ", '" +
                                            rdWL.GetString(0).Trim() + " " + sOffDutyTime + "', " + 
                                            (isHoliday == "Y" ? "'Y'" : "null") + ", '" + sBossEngName + "', '" + sBossLocName + "'," +
                                            (sOutTime == null ? "null" : iSpan.ToString()) + ", " + 
                                            (sOutTime == null ? "null" : iOffSpan.ToString()) + ", " +
                                            (iOver == 0 ? "null" : iOver.ToString()) + ", " + 
                                            (sAllowSupper == null ? "null" : "'" + sAllowSupper.Trim() + "'") + ", " +
                                            (sAllowTaxi == null ? "null" : "'" + sAllowTaxi.Trim() + "'") + ", " +
                                            (sDeptCode == "" ? "null" : "'" + sDeptCode + "'") + ", " + 
                                            (sDeptName == "" ? "null" : "'" + sDeptName + "'") + ", GetDate(), -3, GetDate())";
                                        cmdMIS.ExecuteNonQuery();
                                    }

                                }
                            }
                            catch (Exception ex)
                            {
                                DANCom.DANClass.SendMail(RECIPIENT, "員工出勤資料轉檔有問題！", @"員工出勤資料轉檔有問題\n" + " - " + sEmpNo + " - " + ex.Message);
                                DANCom.DANClass.WriteToLog(PROG_ID, 6, "E", PROG_USER_ID, sComputerName, "第" + i.ToString() + @"筆資料有錯\n" + ex.Message);
                            }
                        }
                        rdWL.Close();

                        //DANCom.DANClass.WriteToLog(PROG_ID, 6, "D", PROG_USER_ID, sComputerName, "員工出勤資料轉檔結束");
                    }
                    catch (Exception ex)
                    {
                        DANCom.DANClass.SendMail(RECIPIENT, "員工出勤資料轉檔錯誤！", "員工出勤資料轉檔錯誤" + " - " + sDBType + " - " + sEmpNo + " - " + ex.Message);
                        DANCom.DANClass.WriteToLog(PROG_ID, 6, "E", PROG_USER_ID, sComputerName, sDBType + " - " + ex.Message);
                    }
                }
                #endregion

                #region 從補刷卡流程檔之轉檔 2017/11/23新增
                sDBType = "Wiselan-2";
                using (SqlConnection cnWL = new SqlConnection(strWiselan))
                {
                    string sEmpNo = "";
                    DataTable dtApp = new DataTable();

                    try
                    {
                        cnWL.Open();
                        //若補刷卡流程在這段時間內有主管已經核過的，則再重新轉檔
                        SqlDataAdapter ada = new SqlDataAdapter("SELECT empno, wdat FROM Appredflow WHERE step1date >= @dat", cnWL);
                        ada.SelectCommand.Parameters.AddWithValue("@dat", DateTime.Today.AddDays(-7).ToString("yyyy/MM/dd"));
                        ada.Fill(dtApp);

                        if (dtApp.Rows.Count > 0)
                        {
                            for (int rowCount = 0; rowCount < dtApp.Rows.Count; rowCount++)
                            {
                                SqlCommand cmdWL = new SqlCommand("SELECT wdat, empno, time1, time2, wtime1, wtime2, daytyp FROM History WHERE empno = @no AND wdat = @dat", cnWL);
                                cmdWL.Parameters.AddWithValue("@no", dtApp.Rows[rowCount]["empno"].ToString());
                                cmdWL.Parameters.AddWithValue("@dat", dtApp.Rows[rowCount]["wdat"].ToString());
                                SqlDataReader rdWL = cmdWL.ExecuteReader();
                                SqlConnection cnOff = new SqlConnection(strWiselan);
                                cnOff.Open();
                                SqlCommand cmdOff = new SqlCommand();
                                cmdOff.Connection = cnOff;
                                SqlDataReader rdOff;

                                if (rdWL.HasRows)
                                {
                                    int i = 0;

                                    try
                                    {
                                        //string abc;
                                        const int NOONSPAN = 60;
                                        string sInTime;
                                        string sOutTime;
                                        int iOffSpan = 0;
                                        int iSpan = 0;
                                        int iOver = 0;
                                        string isHoliday;

                                        SqlConnection cnMIS = new SqlConnection(strGCMIS);
                                        cnMIS.Open();

                                        SqlCommand cmdMIS = new SqlCommand();
                                        cmdMIS.Connection = cnMIS;
                                        SqlDataReader rdMIS;

                                        string sCompCode;
                                        string sEmpEngName;
                                        string sEmpLocName;
                                        Int32 iEmpUniqNo;
                                        string sBossEngName;
                                        string sBossLocName;
                                        string sDeptCode;
                                        string sDeptName;
                                        string sOfficeCode;
                                        string sOnDutyTime;
                                        string sOffDutyTime;

                                        while (rdWL.Read())
                                        {
                                            sDBType = "MIS";
                                            //2015/11/02 Robert 新增
                                            //Isobar 吳宛蓉(Maggie Wu) 2015/10/31離職，但繼續留下來支援，故其出勤資料予以忽略
                                            //if (rdWL.GetString(1).Trim() == "8686")
                                            //    continue;
                                            //if (rdWL.GetString(1).Trim() == "9739")
                                            //    sDBType = "MIS";

                                            //2016/05/05 Robert 新增
                                            //因為有些兼職員工轉正職後，卡機資料不會更新，因此會對應不到新的員工編號，如此更新後就不會再顯示錯誤訊息
                                            if (!CheckEmployee(rdWL.GetString(1), rdWL.GetString(0))) continue;

                                            sEmpNo = rdWL.GetString(1).Trim();

                                            i++;
                                            cmdOff.CommandText = "SELECT offtime FROM Offtra WHERE empno = '" + rdWL.GetString(1).Trim() + "' AND wdat = '" + rdWL.GetString(0).Trim() + "'";
                                            rdOff = cmdOff.ExecuteReader();
                                            if (rdOff.HasRows)
                                            {
                                                rdOff.Read();
                                                iOffSpan = (int)rdOff.GetDouble(0) * 60;
                                            }
                                            else
                                                iOffSpan = 0;
                                            rdOff.Close();

                                            cmdOff.CommandText = "SELECT * FROM Holday WHERE wdat = '" + rdWL.GetString(0).Trim() + "'";
                                            rdOff = cmdOff.ExecuteReader();
                                            if (rdOff.HasRows)
                                                isHoliday = "Y";
                                            else
                                                isHoliday = "N";
                                            rdOff.Close();

                                            if (rdWL.GetString(4).Trim() != "")
                                                sInTime = rdWL.GetString(4).Trim().Substring(0, 2) + ":" + rdWL.GetString(4).Trim().Substring(2) + ":00";
                                            else
                                                sInTime = null;

                                            if (rdWL.GetString(5).Trim() != "")
                                            {
                                                if (rdWL.GetString(6).Trim() == "Y")
                                                {
                                                    sOutTime = DateTime.Parse(rdWL.GetString(0).Trim()).AddDays(1).ToString("yyyy/MM/dd") + " " + rdWL.GetString(5).Trim().Substring(0, 2) + ":" + rdWL.GetString(5).Trim().Substring(2) + ":00";
                                                }
                                                else
                                                    sOutTime = rdWL.GetString(0).Trim() + " " + rdWL.GetString(5).Trim().Substring(0, 2) + ":" + rdWL.GetString(5).Trim().Substring(2) + ":00";

                                                if (rdWL.GetString(4).Trim() != "")
                                                    if (isHoliday == "N")
                                                        iSpan = (int)(DateTime.Parse(sOutTime) - DateTime.Parse(rdWL.GetString(0).Trim() + " " + sInTime)).TotalMinutes + iOffSpan - NOONSPAN;
                                                    else
                                                        iSpan = (int)(DateTime.Parse(sOutTime) - DateTime.Parse(rdWL.GetString(0).Trim() + " " + sInTime)).TotalMinutes + iOffSpan;
                                                else
                                                    iSpan = 0;
                                            }
                                            else
                                            {
                                                sOutTime = null;
                                                iSpan = 0;
                                            }

                                            if (iSpan > 480)
                                                iOver = iSpan - 480;
                                            else
                                                iOver = 0;

                                            string sAllowSupper = null;
                                            string sAllowTaxi = null;

                                            if (isHoliday == "Y")   //假日
                                            {
                                                if (iSpan >= 480)
                                                    sAllowSupper = "YES(2)";
                                                else
                                                {
                                                    if (iSpan >= 240)
                                                        sAllowSupper = "YES";
                                                }

                                                if (iSpan >= 360)
                                                {
                                                    sAllowTaxi = "YES";
                                                }
                                            }
                                            else                    //上班日
                                            {
                                                if (iSpan >= 600 && (sOutTime.CompareTo(rdWL.GetString(0).Trim() + " 20:00:00") >= 0))
                                                    sAllowSupper = "YES";
                                                else
                                                    sAllowSupper = null;

                                                if (iSpan >= 720 && (sOutTime.CompareTo(rdWL.GetString(0).Trim() + " 22:00:00") >= 0))
                                                    sAllowTaxi = "YES";
                                                else
                                                    sAllowTaxi = null;
                                            }

                                            getEmpData(rdWL.GetString(1).Trim(), rdWL.GetString(0).Trim(), out sEmpEngName, out sEmpLocName, out sCompCode, out iEmpUniqNo, out sBossEngName, out sBossLocName, out sDeptCode, out sDeptName, out sOfficeCode);
                                            GetDutyTime(iEmpUniqNo, out sOnDutyTime, out sOffDutyTime);
                                            cmdMIS.CommandText = "SELECT ISNULL(AllowTaxi, ''), ISNULL(AllowSupper, ''), ISNULL(ConfirmInTime, ''), LastCardTime, CheckOutTime, ConfirmOutTime, ISNULL(TotalWorkTime, 0), ISNULL(OverTime, 0) FROM EmpAttendance WHERE RegionCode = 'TW' AND WorkDate = '" + rdWL.GetString(0).Trim() + "' AND EmpNo = '" + rdWL.GetString(1).Trim() + "' AND CompCode = '" + sCompCode + "' AND OfficeCode = '" + sOfficeCode + "' AND EmpUniqNo = " + iEmpUniqNo.ToString();
                                            rdMIS = cmdMIS.ExecuteReader();
                                            if (rdMIS.HasRows)
                                            {
                                                rdMIS.Read();

                                                sAllowTaxi = (rdMIS.GetString(0).Trim() == "" ? sAllowTaxi : rdMIS.GetString(0).Trim());
                                                sAllowSupper = (rdMIS.GetString(1).Trim() == "" ? sAllowSupper : rdMIS.GetString(1).Trim());

                                                if (rdMIS.GetString(0).Trim().CompareTo(sAllowTaxi == null ? "" : sAllowTaxi) != 0 ||
                                                    rdMIS.GetString(1).Trim().CompareTo(sAllowSupper == null ? "" : sAllowSupper) != 0 ||
                                                    rdMIS.GetString(2).Trim().CompareTo(sInTime) != 0 ||
                                                    (rdMIS.IsDBNull(3) ? "9999/12/31 23:59:59" : rdMIS.GetDateTime(3).ToString("yyyy/MM/dd HH:mm:ss")) != (sOutTime == null ? "9999/12/31 23:59:59" : sOutTime) ||
                                                    (rdMIS.IsDBNull(5) ? "9999/12/31 23:59:59" : rdMIS.GetDateTime(5).ToString("yyyy/MM/dd HH:mm:ss")) != (sOutTime == null ? "9999/12/31 23:59:59" : sOutTime) ||
                                                    rdMIS.GetInt32(6) != iSpan || rdMIS.GetInt32(7) != iOver)
                                                {
                                                    rdMIS.Close();

                                                    cmdMIS.CommandText = "UPDATE EmpAttendance SET RegionCode = 'TW'," +
                                                        " CompCode = '" + sCompCode + "', OfficeCode = 'TPE', " +
                                                        "EmpUniqNo = " + iEmpUniqNo.ToString() +
                                                        ", WorkDate = '" + rdWL.GetString(0).Trim() +
                                                        "', EmpNo = '" + rdWL.GetString(1).Trim() +
                                                        "', EmpEngName = '" + sEmpEngName +
                                                        "', EmpLocName = N'" + sEmpLocName +
                                                        "', FirstCardTime = " + (rdWL.GetString(4).Trim() == "" ? "null" : "'" + sInTime + "'") +
                                                        ", InTime = " + (rdWL.GetString(4).Trim() == "" ? "null" : "'" + sInTime + "'") +
                                                        ", ConfirmInTime = " + (rdWL.GetString(4).Trim() == "" ? "null" : "'" + sInTime + "'") +
                                                        ", WorkBegTime = '" + sOnDutyTime +
                                                        "', LastCardTime = " + (sOutTime == null ? "null" : "'" + sOutTime + "'") +
                                                        ", ConfirmOutTime = " + (sOutTime == null ? "null" : "'" + sOutTime + "'") +
                                                        ", WorkOffTime = '" + rdWL.GetString(0).Trim() + " " + sOffDutyTime +
                                                        "', IsHoliday = " + (isHoliday == "Y" ? "'Y'" : "null") +
                                                        ", BossEngName = " + (sBossEngName == "" ? "null" : "'" + sBossEngName + "'") +
                                                        ", BossLocName = " + (sBossLocName == "" ? "null" : "'" + sBossLocName + "'") +
                                                        ", TotalWorkTime = " + (sOutTime == null ? "null" : iSpan.ToString()) +
                                                        ", LeaveSpanTime = " + (sOutTime == null ? "null" : iOffSpan.ToString()) +
                                                        ", OverTime = " + (iOver == 0 ? "null" : iOver.ToString()) +
                                                        ", AllowSupper = " + (sAllowSupper == null ? "null" : "'" + sAllowSupper.Trim() + "'") +
                                                        ", AllowTaxi = " + (sAllowTaxi == null ? "null" : "'" + sAllowTaxi.Trim() + "'") +
                                                        ", DeptCode = " + (sDeptCode == "" ? "null" : "'" + sDeptCode + "'") +
                                                        ", DeptName = " + (sDeptName == "" ? "null" : "'" + sDeptName + "'") +
                                                        ", ChangeBy = -2, ChangeDate = getdate() WHERE RegionCode = 'TW' AND WorkDate = '" +
                                                        rdWL.GetString(0).Trim() + "' AND CompCode = '" + sCompCode + "' AND OfficeCode = 'TPE' AND EmpUniqNo = " + iEmpUniqNo.ToString();
                                                    cmdMIS.ExecuteNonQuery();
                                                }
                                                else
                                                {
                                                    rdMIS.Close();
                                                }
                                            }
                                            else
                                            {
                                                rdMIS.Close();

                                                getEmpData(rdWL.GetString(1).Trim(), rdWL.GetString(0).Trim(), out sEmpEngName, out sEmpLocName, out sCompCode, out iEmpUniqNo, out sBossEngName, out sBossLocName, out sDeptCode, out sDeptName, out sOfficeCode);

                                                cmdMIS.CommandText = "INSERT INTO EmpAttendance (RegionCode, CompCode, " +
                                                    "OfficeCode, EmpUniqNo, WorkDate, EmpNo, EmpEngName, EmpLocName, " +
                                                    "FirstCardTime, InTime, ConfirmInTime, WorkBegTime, LastCardTime, ConfirmOutTime, WorkOffTime, IsHoliday, " +
                                                    "BossEngName, BossLocName, TotalWorkTime, LeaveSpanTime, " +
                                                    "OverTime, AllowSupper, AllowTaxi, DeptCode, DeptName, InputDate, ChangeBy, ChangeDate) " +
                                                    " VALUES ('TW', '" +
                                                    sCompCode + "', 'TPE', " + iEmpUniqNo.ToString() + ", '" +
                                                    rdWL.GetString(0).Trim() + "', '" + rdWL.GetString(1).Trim() + "', '" +
                                                    sEmpEngName + "', N'" + sEmpLocName + "', " +
                                                    (rdWL.GetString(4).Trim() == "" ? "null" : "'" + sInTime + "'") + ", " +
                                                    (rdWL.GetString(4).Trim() == "" ? "null" : "'" + sInTime + "'") + ", " +
                                                    (rdWL.GetString(4).Trim() == "" ? "null" : "'" + sInTime + "'") + ", '" +
                                                    sOnDutyTime + "', " + (sOutTime == null ? "null" : "'" + sOutTime + "'") + ", " +
                                                    (sOutTime == null ? "null" : "'" + sOutTime + "'") + ", '" +
                                                    rdWL.GetString(0).Trim() + " " + sOffDutyTime + "', " +
                                                    (isHoliday == "Y" ? "'Y'" : "null") + ", '" + sBossEngName + "', '" + sBossLocName + "'," +
                                                    (sOutTime == null ? "null" : iSpan.ToString()) + ", " +
                                                    (sOutTime == null ? "null" : iOffSpan.ToString()) + ", " +
                                                    (iOver == 0 ? "null" : iOver.ToString()) + ", " +
                                                    (sAllowSupper == null ? "null" : "'" + sAllowSupper.Trim() + "'") + ", " +
                                                    (sAllowTaxi == null ? "null" : "'" + sAllowTaxi.Trim() + "'") + ", " +
                                                    (sDeptCode == "" ? "null" : "'" + sDeptCode + "'") + ", " +
                                                    (sDeptName == "" ? "null" : "'" + sDeptName + "'") + ", GetDate(), -3, GetDate())";
                                                cmdMIS.ExecuteNonQuery();
                                            }

                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        DANCom.DANClass.SendMail(RECIPIENT, "員工出勤資料轉檔有問題！", @"員工出勤資料轉檔有問題\n" + " - " + sEmpNo + " - " + ex.Message);
                                        DANCom.DANClass.WriteToLog(PROG_ID, 6, "E", PROG_USER_ID, sComputerName, "第" + i.ToString() + @"筆資料有錯\n" + ex.Message);
                                    }
                                }
                                rdWL.Close();
                            }
                        }
                        

                        //DANCom.DANClass.WriteToLog(PROG_ID, 6, "D", PROG_USER_ID, sComputerName, "員工出勤資料轉檔結束");
                    }
                    catch (Exception ex)
                    {
                        DANCom.DANClass.SendMail(RECIPIENT, "員工出勤資料轉檔錯誤！", "員工出勤資料轉檔錯誤" + " - " + sDBType + " - " + sEmpNo + " - " + ex.Message);
                        DANCom.DANClass.WriteToLog(PROG_ID, 6, "E", PROG_USER_ID, sComputerName, sDBType + " - " + ex.Message);
                    }
                }
                #endregion
            }
            #endregion

            #region 轉員工出勤補刷資料
            if (bFill)
            {
                //DANCom.DANClass.WriteToLog(PROG_ID, 6, "D", PROG_USER_ID, sComputerName, "開始轉員工出勤補刷資料");

                sDBType = "Wiselan";
                using (SqlConnection cnWL = new SqlConnection(strWiselan))
                {
                    string sEmpNo = "";

                    try
                    {
                        cnWL.Open();

                        DataTable dtFill = new DataTable();

                        SqlDataAdapter ada = new SqlDataAdapter("SELECT empno, wdat FROM Appredflow WHERE step1date >= '" + DateTime.Today.AddDays(-2).ToString("yyyy/MM/dd") + "'", cnWL);
                        //SqlDataAdapter ada = new SqlDataAdapter("SELECT empno, wdat FROM Appredflow WHERE step1date >= '2017/01/25'", cnWL);
                        ada.Fill(dtFill);

                        if (dtFill.Rows.Count > 0)
                        {
                            SqlCommand cmdWL = new SqlCommand();
                            cmdWL.Connection = cnWL;
                            SqlDataReader rdWL;
                            SqlConnection cnOff = new SqlConnection(strWiselan);
                            cnOff.Open();
                            SqlCommand cmdOff = new SqlCommand();
                            cmdOff.Connection = cnOff;
                            SqlDataReader rdOff;

                            for (int iFill = 0; iFill < dtFill.Rows.Count; iFill++)
                            {
                                //if (dtFill.Rows[iFill]["empno"].ToString() == "9276")
                                //    cmdWL.CommandText = "SELECT wdat, empno, time1, time2, wtime1, wtime2, daytyp FROM History WHERE wdat = @date AND empno = @no";
                                cmdWL.CommandText = "SELECT wdat, empno, time1, time2, wtime1, wtime2, daytyp FROM History WHERE wdat = @date AND empno = @no";
                                cmdWL.Parameters.Clear();
                                cmdWL.Parameters.Add(new SqlParameter("@date", SqlDbType.VarChar));
                                cmdWL.Parameters.Add(new SqlParameter("@no", SqlDbType.VarChar));
                                cmdWL.Parameters["@date"].Value = dtFill.Rows[iFill]["wdat"].ToString();
                                cmdWL.Parameters["@no"].Value = dtFill.Rows[iFill]["empno"].ToString();
                                rdWL = cmdWL.ExecuteReader();
                                
                                if (rdWL.HasRows)
                                {
                                    int i = 0;

                                    try
                                    {
                                        const int NOONSPAN = 60;
                                        string sInTime;
                                        string sOutTime;
                                        int iOffSpan = 0;
                                        int iSpan = 0;
                                        int iOver = 0;
                                        string isHoliday;

                                        SqlConnection cnMIS = new SqlConnection(strGCMIS);
                                        cnMIS.Open();

                                        SqlCommand cmdMIS = new SqlCommand();
                                        cmdMIS.Connection = cnMIS;
                                        SqlDataReader rdMIS;

                                        string sCompCode;
                                        string sEmpEngName;
                                        string sEmpLocName;
                                        Int32 iEmpUniqNo;
                                        string sBossEngName;
                                        string sBossLocName;
                                        string sDeptCode;
                                        string sDeptName;
                                        string sOfficeCode;
                                        string sOnDutyTime;
                                        string sOffDutyTime;

                                        rdWL.Read();
                                        sDBType = "MIS";
                                        sEmpNo = rdWL.GetString(1).Trim();
                                        //if (sEmpNo == "DKE0831")
                                        //    sDBType = "MIS";

                                        i++;
                                        cmdOff.CommandText = "SELECT offtime FROM Offtra WHERE empno = '" + rdWL.GetString(1).Trim() + "' AND wdat = '" + rdWL.GetString(0).Trim() + "'";
                                        rdOff = cmdOff.ExecuteReader();
                                        if (rdOff.HasRows)
                                        {
                                            rdOff.Read();
                                            iOffSpan = (int)rdOff.GetDouble(0) * 60;
                                        }
                                        else
                                            iOffSpan = 0;
                                        rdOff.Close();

                                        cmdOff.CommandText = "SELECT * FROM Holday WHERE wdat = '" + rdWL.GetString(0).Trim() + "'";
                                        rdOff = cmdOff.ExecuteReader();
                                        if (rdOff.HasRows)
                                            isHoliday = "Y";
                                        else
                                            isHoliday = "N";
                                        rdOff.Close();

                                        if (rdWL.GetString(4).Trim() != "")
                                            sInTime = rdWL.GetString(4).Trim().Substring(0, 2) + ":" + rdWL.GetString(4).Trim().Substring(2) + ":00";
                                        else
                                            sInTime = null;

                                        if (rdWL.GetString(5).Trim() != "")
                                        {
                                            if (rdWL.GetString(6).Trim() == "Y")
                                            {
                                                sOutTime = DateTime.Parse(rdWL.GetString(0).Trim()).AddDays(1).ToString("yyyy/MM/dd") + " " + rdWL.GetString(5).Trim().Substring(0, 2) + ":" + rdWL.GetString(5).Trim().Substring(2) + ":00";
                                            }
                                            else
                                                sOutTime = rdWL.GetString(0).Trim() + " " + rdWL.GetString(5).Trim().Substring(0, 2) + ":" + rdWL.GetString(5).Trim().Substring(2) + ":00";

                                            if (rdWL.GetString(4).Trim() != "")
                                                if (isHoliday == "N")
                                                    iSpan = (int)(DateTime.Parse(sOutTime) - DateTime.Parse(rdWL.GetString(0).Trim() + " " + sInTime)).TotalMinutes + iOffSpan - NOONSPAN;
                                                else
                                                    iSpan = (int)(DateTime.Parse(sOutTime) - DateTime.Parse(rdWL.GetString(0).Trim() + " " + sInTime)).TotalMinutes + iOffSpan;
                                            else
                                                iSpan = 0;
                                        }
                                        else
                                        {
                                            sOutTime = null;
                                            iSpan = 0;
                                        }

                                        if (iSpan > 480)
                                            iOver = iSpan - 480;
                                        else
                                            iOver = 0;

                                        string sAllowSupper = null;
                                        string sAllowTaxi = null;

                                        if (isHoliday == "Y")   //假日
                                        {
                                            if (iSpan >= 480)
                                                sAllowSupper = "YES(2)";
                                            else
                                            {
                                                if (iSpan >= 240)
                                                    sAllowSupper = "YES";
                                            }

                                            if (iSpan >= 360)
                                            {
                                                sAllowTaxi = "YES";
                                            }
                                        }
                                        else                    //上班日
                                        {
                                            if (iSpan >= 600 && (sOutTime.CompareTo(rdWL.GetString(0).Trim() + " 20:00:00") >= 0))
                                                sAllowSupper = "YES";
                                            else
                                                sAllowSupper = null;

                                            if (iSpan >= 720 && (sOutTime.CompareTo(rdWL.GetString(0).Trim() + " 22:00:00") >= 0))
                                                sAllowTaxi = "YES";
                                            else
                                                sAllowTaxi = null;
                                        }

                                        getEmpData(rdWL.GetString(1).Trim(), rdWL.GetString(0).Trim(), out sEmpEngName, out sEmpLocName, out sCompCode, out iEmpUniqNo, out sBossEngName, out sBossLocName, out sDeptCode, out sDeptName, out sOfficeCode);
                                        GetDutyTime(iEmpUniqNo, out sOnDutyTime, out sOffDutyTime);
                                        cmdMIS.CommandText = "SELECT ISNULL(AllowTaxi, ''), ISNULL(AllowSupper, ''), ISNULL(ConfirmInTime, ''), LastCardTime, CheckOutTime, ConfirmOutTime, ISNULL(TotalWorkTime, 0), ISNULL(OverTime, 0) FROM EmpAttendance WHERE RegionCode = 'TW' AND WorkDate = '" + rdWL.GetString(0).Trim() + "' AND EmpNo = '" + rdWL.GetString(1).Trim() + "' AND CompCode = '" + sCompCode + "' AND OfficeCode = '" + sOfficeCode + "' AND EmpUniqNo = " + iEmpUniqNo.ToString();
                                        rdMIS = cmdMIS.ExecuteReader();
                                        if (rdMIS.HasRows)
                                        {
                                            rdMIS.Read();

                                            sAllowTaxi = (rdMIS.GetString(0).Trim() == "" ? sAllowTaxi : rdMIS.GetString(0).Trim());
                                            sAllowSupper = (rdMIS.GetString(1).Trim() == "" ? sAllowSupper : rdMIS.GetString(1).Trim());

                                            if (rdMIS.GetString(0).Trim().CompareTo(sAllowTaxi == null ? "" : sAllowTaxi) != 0 ||
                                                rdMIS.GetString(1).Trim().CompareTo(sAllowSupper == null ? "" : sAllowSupper) != 0 ||
                                                rdMIS.GetString(2).Trim().CompareTo(sInTime) != 0 ||
                                                (rdMIS.IsDBNull(3) ? "9999/12/31 23:59:59" : rdMIS.GetDateTime(3).ToString("yyyy/MM/dd HH:mm:ss")) != (sOutTime == null ? "9999/12/31 23:59:59" : sOutTime) ||
                                                (rdMIS.IsDBNull(5) ? "9999/12/31 23:59:59" : rdMIS.GetDateTime(5).ToString("yyyy/MM/dd HH:mm:ss")) != (sOutTime == null ? "9999/12/31 23:59:59" : sOutTime) ||
                                                rdMIS.GetInt32(6) != iSpan || rdMIS.GetInt32(7) != iOver)
                                            {
                                                rdMIS.Close();



                                                cmdMIS.CommandText = "UPDATE EmpAttendance SET RegionCode = 'TW'," +
                                                    " CompCode = '" + sCompCode + "', OfficeCode = 'TPE', " +
                                                    "EmpUniqNo = " + iEmpUniqNo.ToString() +
                                                    ", WorkDate = '" + rdWL.GetString(0).Trim() +
                                                    "', EmpNo = '" + rdWL.GetString(1).Trim() +
                                                    "', EmpEngName = '" + sEmpEngName +
                                                    "', EmpLocName = '" + sEmpLocName +
                                                    "', FirstCardTime = " + (rdWL.GetString(4).Trim() == "" ? "null" : "'" + sInTime + "'") +
                                                    ", InTime = " + (rdWL.GetString(4).Trim() == "" ? "null" : "'" + sInTime + "'") +
                                                    ", ConfirmInTime = " + (rdWL.GetString(4).Trim() == "" ? "null" : "'" + sInTime + "'") +
                                                    ", WorkBegTime = '" + sOnDutyTime +
                                                    "', LastCardTime = " + (sOutTime == null ? "null" : "'" + sOutTime + "'") +
                                                    ", ConfirmOutTime = " + (sOutTime == null ? "null" : "'" + sOutTime + "'") +
                                                    ", WorkOffTime = '" + rdWL.GetString(0).Trim() + " " + sOffDutyTime +
                                                    "', IsHoliday = " + (isHoliday == "Y" ? "'Y'" : "null") +
                                                    ", BossEngName = " + (sBossEngName == "" ? "null" : "'" + sBossEngName + "'") +
                                                    ", BossLocName = " + (sBossLocName == "" ? "null" : "'" + sBossLocName + "'") +
                                                    ", TotalWorkTime = " + (sOutTime == null ? "null" : iSpan.ToString()) +
                                                    ", LeaveSpanTime = " + (sOutTime == null ? "null" : iOffSpan.ToString()) +
                                                    ", OverTime = " + (iOver == 0 ? "null" : iOver.ToString()) +
                                                    ", AllowSupper = " + (sAllowSupper == null ? "null" : "'" + sAllowSupper.Trim() + "'") +
                                                    ", AllowTaxi = " + (sAllowTaxi == null ? "null" : "'" + sAllowTaxi.Trim() + "'") +
                                                    ", DeptCode = " + (sDeptCode == "" ? "null" : "'" + sDeptCode + "'") +
                                                    ", DeptName = " + (sDeptName == "" ? "null" : "'" + sDeptName + "'") +
                                                    ", ChangeBy = -4, ChangeDate = getdate() WHERE RegionCode = 'TW' AND WorkDate = '" +
                                                    rdWL.GetString(0).Trim() + "' AND CompCode = '" + sCompCode + "' AND OfficeCode = 'TPE' AND EmpUniqNo = " + iEmpUniqNo.ToString();
                                                cmdMIS.ExecuteNonQuery();
                                            }
                                            else
                                            {
                                                rdMIS.Close();
                                            }
                                        }
                                        else
                                        {
                                            rdMIS.Close();

                                            getEmpData(rdWL.GetString(1).Trim(), rdWL.GetString(0).Trim(), out sEmpEngName, out sEmpLocName, out sCompCode, out iEmpUniqNo, out sBossEngName, out sBossLocName, out sDeptCode, out sDeptName, out sOfficeCode);

                                            cmdMIS.CommandText = "INSERT INTO EmpAttendance (RegionCode, CompCode, " +
                                                "OfficeCode, EmpUniqNo, WorkDate, EmpNo, EmpEngName, EmpLocName, " +
                                                "FirstCardTime, InTime, ConfirmInTime, WorkBegTime, LastCardTime, ConfirmOutTime, WorkOffTime, IsHoliday, " +
                                                "BossEngName, BossLocName, TotalWorkTime, LeaveSpanTime, " +
                                                "OverTime, AllowSupper, AllowTaxi, DeptCode, DeptName, InputDate, ChangeBy, ChangeDate) " +
                                                " VALUES ('TW', '" +
                                                sCompCode + "', 'TPE', " + iEmpUniqNo.ToString() + ", '" +
                                                rdWL.GetString(0).Trim() + "', '" + rdWL.GetString(1).Trim() + "', '" +
                                                sEmpEngName + "', '" + sEmpLocName + "', " +
                                                (rdWL.GetString(4).Trim() == "" ? "null" : "'" + sInTime + "'") + ", " +
                                                (rdWL.GetString(4).Trim() == "" ? "null" : "'" + sInTime + "'") + ", " +
                                                (rdWL.GetString(4).Trim() == "" ? "null" : "'" + sInTime + "'") + ", '" +
                                                sOnDutyTime + "', " + (sOutTime == null ? "null" : "'" + sOutTime + "'") + ", " +
                                                (sOutTime == null ? "null" : "'" + sOutTime + "'") + ", '" +
                                                rdWL.GetString(0).Trim() + " " + sOffDutyTime + "', " +
                                                (isHoliday == "Y" ? "'Y'" : "null") + ", '" + sBossEngName + "', '" + sBossLocName + "'," +
                                                (sOutTime == null ? "null" : iSpan.ToString()) + ", " +
                                                (sOutTime == null ? "null" : iOffSpan.ToString()) + ", " +
                                                (iOver == 0 ? "null" : iOver.ToString()) + ", " +
                                                (sAllowSupper == null ? "null" : "'" + sAllowSupper.Trim() + "'") + ", " +
                                                (sAllowTaxi == null ? "null" : "'" + sAllowTaxi.Trim() + "'") + ", " +
                                                (sDeptCode == "" ? "null" : "'" + sDeptCode + "'") + ", " +
                                                (sDeptName == "" ? "null" : "'" + sDeptName + "'") + ", GetDate(), -5, GetDate())";
                                            cmdMIS.ExecuteNonQuery();
                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        DANCom.DANClass.SendMail(RECIPIENT, "員工出勤補刷資料轉檔有問題！", @"員工出勤補刷資料轉檔有問題\n" + " - " + sEmpNo + " - " + ex.Message);
                                        DANCom.DANClass.WriteToLog(PROG_ID, 6, "E", PROG_USER_ID, sComputerName, "第" + i.ToString() + @"筆資料有錯\n" + ex.Message);
                                    }
                                }
                                rdWL.Close();
                            }
                        }

                        //DANCom.DANClass.WriteToLog(PROG_ID, 6, "D", PROG_USER_ID, sComputerName, "員工出勤補刷資料轉檔結束");
                    }
                    catch (Exception ex)
                    {
                        DANCom.DANClass.SendMail(RECIPIENT, "員工出勤補刷資料轉檔錯誤！", "員工出勤補刷資料轉檔錯誤" + " - " + sDBType + " - " + sEmpNo + " - " + ex.Message);
                        DANCom.DANClass.WriteToLog(PROG_ID, 6, "E", PROG_USER_ID, sComputerName, sDBType + " - " + ex.Message);
                    }
                }
            }
            #endregion

        }



        #region 取得員工資料
        static void getEmpData(string sEmpNo, string sWorkDate, out string sEmpEngName, out string sEmpLocName, out string sCompCode,
                out Int32 iEmpUniqNo, out string sBossEngName, out string sBossLocName, out string sDeptCode, 
                out string sDeptName, out string sOfficeCode)
        {
            Int32 iBossUniqNo;

            SqlConnection cnGCMIS = new SqlConnection(strCommon_L);
            SqlCommand cmdGCMIS = new SqlCommand();
            cnGCMIS.Open();
            cmdGCMIS.Connection = cnGCMIS;
            cmdGCMIS.CommandText = "SELECT EmpUniqNo, EmpEngName, EmpLocName, CompCode, BossEmpUniqNo, DeptCode, OfficeCode FROM Employee WHERE RegionCode = 'TW' AND EmpNo = '" + sEmpNo + "' AND (EmpQuitDate IS NULL OR EmpQuitDate = '' OR EmpQuitDate >= '" + sWorkDate + "')";
            SqlDataReader rd = cmdGCMIS.ExecuteReader();
            if (rd.HasRows)
            {
                rd.Read();
                iEmpUniqNo = rd.IsDBNull(0) ? 0 : rd.GetInt32(0);
                sEmpEngName = rd.IsDBNull(1) ? "" : rd.GetString(1);
                sEmpLocName = rd.IsDBNull(2) ? "" : rd.GetString(2);
                sCompCode = rd.IsDBNull(3) ? "" : rd.GetString(3);
                iBossUniqNo = (rd.IsDBNull(4) ? 0 : rd.GetInt32(4));
                sDeptCode = rd.IsDBNull(5) ? "" : rd.GetString(5);
                sOfficeCode = rd.IsDBNull(6) ? "TPE" :rd.GetString(6);
            }
            else
            {
                sEmpEngName = "";
                sEmpLocName = "";
                iEmpUniqNo = 0;
                iBossUniqNo = 0;
                sCompCode = "";
                sDeptCode = "";
                sOfficeCode = "";
            }
            rd.Close();

            cmdGCMIS.CommandText = "SELECT EmpEngName, EmpLocName FROM Employee WHERE EmpUniqNo = " + iBossUniqNo.ToString();
            rd = cmdGCMIS.ExecuteReader();
            if (rd.HasRows)
            {
                rd.Read();
                sBossEngName = rd.GetString(0);
                sBossLocName = rd.GetString(1);
            }
            else
            {
                sBossEngName = "";
                sBossLocName = "";
            }
            rd.Close();

            cmdGCMIS.CommandText = "SELECT DeptLocShortName FROM Dept WHERE RegionCode = 'TW' AND CompCode = '" + sCompCode + "' AND DeptCode = '" + sDeptCode + "'";
            rd = cmdGCMIS.ExecuteReader();
            if (rd.HasRows)
            {
                rd.Read();
                sDeptName = (rd.IsDBNull(0) ? "" : rd.GetString(0));
            }
            else
                sDeptName = "";
            rd.Close();
            cnGCMIS.Close();
        }
        #endregion

        #region 取得標準員工上、下班時間
        static public void GetDutyTime(int iEmpUniqNo, out string sOnDutyTime, out string sOffDutyTime)
        {
            sOnDutyTime = "";
            sOffDutyTime = "";

            using (SqlConnection cn = new SqlConnection(strGCMIS))
            {
                cn.Open();
                SqlCommand cmd = new SqlCommand("SELECT OnDutyTime, OffDutyTime FROM WeekWork WHERE EmpUniqNo = @uniqno", cn);
                cmd.Parameters.Add(new SqlParameter("@uniqno", SqlDbType.Int));
                cmd.Parameters["@uniqno"].Value = iEmpUniqNo;
                SqlDataReader rd = cmd.ExecuteReader();
                if (rd.HasRows)
                {
                    rd.Read();
                    sOnDutyTime = rd.IsDBNull(0) ? "" : rd.GetString(0);
                    sOffDutyTime = rd.IsDBNull(1) ? "" : rd.GetString(1);
                }
                rd.Close();
            }

            if (sOnDutyTime == "" || sOffDutyTime == "")
            {
                string _CompCode = "";

                using (SqlConnection cn = new SqlConnection(strCommon_L))
                {
                    cn.Open();
                    SqlCommand cmd = new SqlCommand("SELECT CompCode FROM Employee WHERE EmpUniqNo = @uniqno", cn);
                    cmd.Parameters.Add(new SqlParameter("@uniqno", SqlDbType.Int));
                    cmd.Parameters["@uniqno"].Value = iEmpUniqNo;
                    SqlDataReader rd = cmd.ExecuteReader();
                    if (rd.HasRows)
                    {
                        rd.Read();
                        _CompCode = rd.GetString(0);
                    }
                    rd.Close();

                    cmd.CommandText = "SELECT WorkLateTime, WorkOffTime FROM OfficeHour WHERE LocCode = @region AND CompCode = @comp";
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add(new SqlParameter("@region", SqlDbType.VarChar));
                    cmd.Parameters.Add(new SqlParameter("@comp", SqlDbType.VarChar));
                    cmd.Parameters["@region"].Value = "TW";
                    cmd.Parameters["@comp"].Value = _CompCode;
                    rd = cmd.ExecuteReader();
                    if (rd.HasRows)
                    {
                        rd.Read();
                        sOnDutyTime = rd.IsDBNull(0) ? "" : rd.GetString(0);
                        sOffDutyTime = rd.IsDBNull(1) ? "" : rd.GetString(1);
                    }
                    rd.Close();
                }
            }
        }
        #endregion

        #region 取得員工總號
        static int GetBossUniqNo(string sEmpNo, string sConnectString)
        {
            int iResult;

            if (sEmpNo == "")
            {
                iResult = 0;
            }
            else
            {
                using (SqlConnection cn = new SqlConnection(sConnectString))
                {
                    cn.Open();
                    SqlCommand cmd = new SqlCommand("SELECT EmpUniqNo FROM Employee WHERE RegionCode = 'TW' AND (EmpQuitDate IS NULL OR EmpQuitDate = '' OR EmpQuitDate >= GetDate()) AND EmpNo = @no", cn);
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add(new SqlParameter("@no", SqlDbType.VarChar));
                    cmd.Parameters["@no"].Value = sEmpNo.Trim();
                    SqlDataReader rd = cmd.ExecuteReader();
                    if (rd.HasRows)
                    {
                        rd.Read();
                        iResult = rd.GetInt32(0);
                    }
                    else
                    {
                        iResult = int.MinValue;
                    }
                    rd.Close();
                }
            }

            return iResult;
        }
        #endregion

        #region 取得 CommonDB 相對的欄位名稱
        static string GetFieldName(string sField)
        {
            string sResult = "";
            if (sField.Trim() == "") return sResult;

            using (SqlConnection cn = new SqlConnection(strCommon_L))
            {
                cn.Open();
                SqlCommand cmd = new SqlCommand("SELECT GCField FROM WLMapping WHERE WLField = @wl", cn);
                cmd.Parameters.Clear();
                cmd.Parameters.Add(new SqlParameter("@wl", SqlDbType.VarChar));
                cmd.Parameters["@wl"].Value = sField;
                SqlDataReader rd = cmd.ExecuteReader();
                if (rd.HasRows)
                {
                    rd.Read();
                    sResult = rd.GetString(0);
                }
                rd.Close();
            }

            return sResult;
        }
        #endregion

        #region 從總號取得員工編號
        static string GetEmpNo(int iNo, string sConnectString)
        {
            string sResult = "";
            if (iNo == 0) return sResult;

            using (SqlConnection cn = new SqlConnection(sConnectString))
            {
                cn.Open();
                SqlCommand cmd = new SqlCommand("SELECT EmpNo FROM Employee WHERE EmpUniqNo = @uniq", cn);
                cmd.Parameters.Clear();
                cmd.Parameters.Add(new SqlParameter("@uniq", SqlDbType.Int));
                cmd.Parameters["@uniq"].Value = iNo;
                SqlDataReader rd = cmd.ExecuteReader();
                if (rd.HasRows)
                {
                    rd.Read();
                    sResult = rd.GetString(0);
                }
                rd.Close();
            }

            return sResult;
        }
        #endregion

        #region 設定新員工密碼
        static void SetEmpPassword(Decimal iNo, string sEmpNo)
        {
            if (iNo == 0 || sEmpNo.Trim() == "") return;

            string sBirthDay = "";
            string sYY = "";
            string sMM = "";
            string sDD = "";

            using (SqlConnection cn = new SqlConnection(strWiselan))
            {
                cn.Open();
                SqlCommand cmd = new SqlCommand("SELECT birt FROM Empmsf WHERE empno = @no", cn);
                cmd.Parameters.Clear();
                cmd.Parameters.Add(new SqlParameter("@no", SqlDbType.VarChar));
                cmd.Parameters["@no"].Value = sEmpNo;
                SqlDataReader rd = cmd.ExecuteReader();
                if (rd.HasRows)
                {
                    rd.Read();
                    sBirthDay = rd.GetString(0);
                }
                rd.Close();
            }

            if (sBirthDay == "") return;

            sYY = (int.Parse(sBirthDay.Substring(0, 4)) - 1911).ToString();
            sMM = sBirthDay.Substring(5, 2);
            sDD = sBirthDay.Substring(8, 2);

            using (SqlConnection cn = new SqlConnection(strCommon_L))
            {
                cn.Open();
                SqlCommand cmd = new SqlCommand("INSERT INTO UserS VALUES (@uniq, @first, @chguser, @chgdate)", cn);
                cmd.Parameters.Clear();
                cmd.Parameters.Add(new SqlParameter("@uniq", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@first", SqlDbType.VarChar));
                cmd.Parameters.Add(new SqlParameter("@chguser", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@chgdate", SqlDbType.DateTime));
                cmd.Parameters["@uniq"].Value = (Int32)iNo;
                if (sEmpNo.Length == 4)
                {
                    cmd.Parameters["@first"].Value = sEmpNo + GetRandom().Substring(0, 4);
                }
                else if (sEmpNo.Length == 5)
                {
                    cmd.Parameters["@first"].Value = sEmpNo.Substring(0, 4) + GetRandom().Substring(0, 4);
                }
                else if (sEmpNo.Length == 6)
                {
                    cmd.Parameters["@first"].Value = sEmpNo.Substring(1, 4) + GetRandom().Substring(0, 4);
                }
                else if (sEmpNo.Length == 7)
                {
                    cmd.Parameters["@first"].Value = sEmpNo.Substring(2, 4) + GetRandom().Substring(0, 4);
                }
                
                cmd.Parameters["@chguser"].Value = 146;
                cmd.Parameters["@chgdate"].Value = DateTime.Now;
                cmd.ExecuteNonQuery();

                cmd.CommandText = "INSERT INTO Partner VALUES (@uniq, @second, @chguser, @chgdate)";
                cmd.Parameters.Clear();
                cmd.Parameters.Add(new SqlParameter("@uniq", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@second", SqlDbType.VarChar));
                cmd.Parameters.Add(new SqlParameter("@chguser", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@chgdate", SqlDbType.DateTime));
                cmd.Parameters["@uniq"].Value = (Int32)iNo;
                cmd.Parameters["@second"].Value = sYY + sMM + sDD;
                cmd.Parameters["@chguser"].Value = 146;
                cmd.Parameters["@chgdate"].Value = DateTime.Now;
                cmd.ExecuteNonQuery();
            }
        }
        #endregion

        #region 取得亂數
        static string GetRandom()
        {
            string sResult = "";

            Random autoRand = new Random();
            sResult = string.Format("{0, 10}", autoRand.Next()).Trim();

            return sResult;
        }
        #endregion

        #region 檢查員工編號是否正確（編號是否存在？員工是否已離職？）
        static bool CheckEmployee(string sEmpNo, string sWorkDate)
        {
            bool bResult = false;

            if (sEmpNo.Trim() == "" || sWorkDate.Trim() == "") return bResult;

            using (SqlConnection cn = new SqlConnection(strCommon_L))
            {
                cn.Open();
                SqlCommand cmd = new SqlCommand("SELECT Count(*) FROM Employee WHERE EmpNo = @no AND (EmpQuitDate IS NULL OR EmpQuitDate = '' OR EmpQuitDate > '" + sWorkDate + "')", cn);
                cmd.Parameters.Clear();
                cmd.Parameters.Add(new SqlParameter("@no", SqlDbType.VarChar));
                cmd.Parameters["@no"].Value = sEmpNo;
                SqlDataReader rd = cmd.ExecuteReader();
                if (rd.HasRows)
                {
                    rd.Read();
                    if (rd.GetInt32(0) > 0) bResult = true;
                }
                rd.Close();
            }

            return bResult;
        }
        #endregion

        #region 取得品牌全名
        static string GetBrandName(string brandShort, string compCode, string deptCode)
        {
            string sResult = "";

            switch (brandShort)
            {
                case "AAA":
                    sResult = "AAA";
                    break;
                case "AMT":
                    sResult = "Amnet";
                    break;
                case "CT":
                    sResult = "Carat";
                    break;
                case "HQ":
                    sResult = "Dentsu Aegis Network";
                    break;
                case "DO":
                    sResult = "Dentsu One";
                    break;
                case "DX":
                    sResult = "Dentsu X";
                    break;
                case "IP":
                    sResult = "iProspect";
                    break;
                case "PO":
                    sResult = "Amplifi";
                    break;
                case "VZ":
                    sResult = "Vizeum";
                    break;
                case "WS":
                    sResult = "Isobar";
                    break;
                case "XL":
                    sResult = "X-Line";
                    break;
                case "MK":
                    sResult = "Merkle";
                    break;
                default:
                    switch (compCode)
                    {
                        case "AAA":
                            sResult = "AAA";
                            break;
                        case "AMT":
                            sResult = "Amnet";
                            break;
                        case "CT":
                            if (deptCode.Substring(0, 1) == "2")
                                sResult = "Dentsu Aegis Network";
                            else
                                sResult = "Carat";
                            break;
                        case "DO":
                            sResult = "Dentsu One";
                            break;
                        case "DX":
                            sResult = "Dentsu X";
                            break;
                        case "IP":
                            sResult = "iProspect";
                            break;
                        case "PO":
                            sResult = "Amplifi";
                            break;
                        case "VZ":
                            sResult = "Vizeum";
                            break;
                        case "WS":
                            sResult = "Isobar";
                            break;
                        case "XL":
                            sResult = "X-Line";
                            break;
                        case "MK":
                            sResult = "Merkle";
                            break;
                        default:
                            sResult = "";
                            break;
                    }
                    break;
            }

            return sResult;
        }
        #endregion

        #region 從JML取得員工EMail帳號
        static string GetEmail(string wdid)
        {
            string sResult = "";

            if (wdid == "") return sResult;
            using (SqlConnection cn = new SqlConnection(strHai))
            {
                cn.Open();
                SqlCommand cmd = new SqlCommand("SELECT EmpCompEmail FROM Employee WHERE WorkdayID = @wdid", cn);
                cmd.Parameters.Clear();
                cmd.Parameters.AddWithValue("@wdid", wdid);
                SqlDataReader rd = cmd.ExecuteReader();
                if (rd.HasRows)
                {
                    rd.Read();
                    sResult = rd.IsDBNull(0) ? "" : rd.GetString(0);
                }
            }

            return sResult;
        }
        #endregion
    }
}
