using Cesco.FW.Global.DBAdapter;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;

namespace DisinfectionAlarmSys
{
    public partial class Usp_DisinfectionAlarmSys : Form
    {

        string _strUserID = string.Empty;
        string _strDeptCode = string.Empty;
        string _strInsertAuth = string.Empty;
        string _strUpdateAuth = string.Empty;
        string _strDeleteAuth = string.Empty;
        string _strSearchAuth = string.Empty;
        string _strPrintAuth = string.Empty;
        string _strExcelAuth = string.Empty;
        string _strDataAuth = string.Empty;

        DataTable Send_DT = new DataTable();
        DataTable EmailList_DT = new DataTable();

        Cesco.FW.Global.DBAdapter.ConfigurationDetail.DBName _dbName;
        Cesco.FW.Global.DBAdapter.DBAdapters dbA;

        string _strDBConnection = string.Empty;
        string _strServerURL = string.Empty;

        public Usp_DisinfectionAlarmSys()
        {
            InitializeComponent();

            //_dbName = Cesco.FW.Global.DBAdapter.ConfigurationDetail.DBName.CESNET2;
            //_strDBConnection = "Data Source=maindb.cesco.biz,11433;Initial Catalog=CESCO_ACCOUNT;Password=cescoeisadm;Persist Security Info=True;User ID=cescoeisadm;";
            _dbName = Cesco.FW.Global.DBAdapter.ConfigurationDetail.DBName.TESTSERVER;
            _strDBConnection = "Data Source=devdb.cesco.biz,11433;Initial Catalog=CESCO_ACCOUNT;Password=15490;Persist Security Info=True;User ID=15490;";

            switch (_dbName.ToString())
            {
                case "TESTSERVER"://TEST서버 연동 작업 안함
                    _strServerURL = "http://cesnetdev.cesco.biz/WCF/IISService/WcfCommon/WcfCommonNew.svc";
                    break;
                case "DEVELOPDB"://TEST서버 연동 작업 안함
                    _strServerURL = "http://cesnetdev.cesco.biz/WCF/IISService/WcfCommon/WcfCommonNew.svc";
                    break;
                default:
                    _strServerURL = "http://cesnet.cesco.biz/WCF/IISService/WcfCommon/WcfCommonNew.svc";
                    break;
            }
        }

        public Usp_DisinfectionAlarmSys(string pUserID, string pDeptCode, string pInsertAuth, string pUpdateAuth, string pDeleteAuth, string pSearchAuth, string pPrintAuth, string pExcelAuth, string pDataAuth)
        {
            InitializeComponent();

            #region CESNET 2.0 전달자
            _strUserID = pUserID;
            _strDeptCode = pDeptCode;
            _strInsertAuth = pInsertAuth;
            _strUpdateAuth = pUpdateAuth;
            _strDeleteAuth = pDeleteAuth;
            _strSearchAuth = pSearchAuth;
            _strPrintAuth = pPrintAuth;
            _strExcelAuth = pExcelAuth;
            _strDataAuth = pDataAuth;
            #endregion
        }

        private void Usp_DisinfectionAlarmSys_Load(object sender, EventArgs e)
        {
            Search();

            this.Dispose();
            this.Close();
        }

        private void Dbconnection()
        {
            try
            {
                dbA = new DBAdapters();
                dbA.LocalInfo = new LocalInfo("99999", System.Reflection.MethodBase.GetCurrentMethod());
                dbA.BindingConfig.ReceiveTimeout = new TimeSpan(0, 40, 30);
                dbA.BindingConfig.SendTimeout = new TimeSpan(0, 40, 30);
                dbA.BindingConfig.CloseTimeout = new TimeSpan(0, 40, 30);
                dbA.BindingConfig.ServerUriString = _strServerURL;
                dbA.BindingConfig.TimeOut = ConfigurationDetail.TimeOuts.MINUTE10;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void Search()
        {
            Dbconnection();

            try
            {
                this.Cursor = Cursors.WaitCursor;

                dbA.Procedure.ProcedureName = "CESNET2.DBO.Usp_Csn_Set_DisInfection_Auto_Alram_Batch_20221202";
                //dbA.Procedure.ProcedureName = "CESNET2.DBO.Usp_Csn_Set_DisInfection_Auto_Alram_Batch";

                DataSet ds = dbA.ProcedureToDataSetCompress();

                // 평일 대상
                if (ds.Tables[0].Rows.Count > 0)
                {
                    DataTable dtMailList = new DataTable();
                    dtMailList = ds.Tables[0];
                    eMailSend(dtMailList, "1");
                }

                // 미전송분 대상
                if (ds.Tables[1].Rows.Count > 0)
                {
                    DataTable dtNotMailList = new DataTable();
                    dtNotMailList = ds.Tables[1];
                    eMailSend(dtNotMailList, "2");
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void eMailSend(DataTable pDtList, string pGubn)
        {
            if (pGubn == "1")
            {
                foreach (DataRow dRow in pDtList.Select())
                {
                    Cesco.FW.Global.Groupware.MailInfos mailInfos = new Cesco.FW.Global.Groupware.MailInfos();

                    mailInfos.Subject = dRow["RTLE"].ToString();
                    mailInfos.IsHtml = true;
                    mailInfos.Body = "<html><body>" +
                                        "" + dRow["RMMO"].ToString() +
                                     "</body></html>";

                    mailInfos.SendEmail = new Cesco.FW.Global.Groupware.EmailAddress("시스템", "postmaster@cesco.co.kr"); // 발신자
                    //mailInfos.ToEmails.Add(new Cesco.FW.Global.Groupware.EmailAddress("", dRow["RNMG"].ToString())); // 수신자
                    mailInfos.ToEmails.Add(new Cesco.FW.Global.Groupware.EmailAddress("", "BH0719@cesco.co.kr")); // 수신자

                    var result = Cesco.FW.Global.Groupware.MailController.Send(mailInfos);

                    // 결과값 처리
                    if (!result.Result)
                    {
                        MessageBox.Show(result.Message);
                    }
                    else
                    {
                        Dbconnection();

                        try
                        {
                            dbA.LocalInfo = new Cesco.FW.Global.DBAdapter.LocalInfo("99999", System.Reflection.MethodBase.GetCurrentMethod());
                            dbA.BindingConfig.ReceiveTimeout = new TimeSpan(0, 0, 40);
                            dbA.BindingConfig.CloseTimeout = new TimeSpan(0, 0, 40);
                            dbA.BindingConfig.TimeOut = ConfigurationDetail.TimeOuts.MINUTE10;
                            dbA.BindingConfig.SendTimeout = new TimeSpan(0, 15, 0);

                            dbA.Procedure.ProcedureName = "CESNET2.dbo.Usp_Csn_Set_DisInfection_History";

                            dbA.Procedure.ParamAdd("@GUBN", pGubn);
                            dbA.Procedure.ParamAdd("@CFID", dRow["CFID"].ToString());
                            dbA.Procedure.ParamAdd("@SSID", dRow["SSID"].ToString());
                            dbA.Procedure.ParamAdd("@EXDT", dRow["EXDT"].ToString());
                            dbA.Procedure.ParamAdd("@RGBN", dRow["RGBN"].ToString());
                            dbA.Procedure.ParamAdd("@RNMG", dRow["RNMG"].ToString());
                            dbA.Procedure.ParamAdd("@ADYN", dRow["ADYN"].ToString());
                            dbA.Procedure.ParamAdd("@RTLE", dRow["RTLE"].ToString());
                            dbA.Procedure.ParamAdd("@RMMO", dRow["RMMO"].ToString());

                            dbA.ProcedureToDataSetCompress();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            this.Cursor = Cursors.Default;
                        }
                    }

                    mailInfos = null;
                }
            }
            else if (pGubn == "2")
            {
                foreach (DataRow dRow in pDtList.Select())
                {
                    Cesco.FW.Global.Groupware.MailInfos mailInfos = new Cesco.FW.Global.Groupware.MailInfos();

                    mailInfos.Subject = dRow["RTLE"].ToString();
                    mailInfos.IsHtml = true;
                    mailInfos.Body = "<html><body>" +
                                        "" + dRow["RMMO"].ToString() +
                                     "</body></html>";

                    mailInfos.SendEmail = new Cesco.FW.Global.Groupware.EmailAddress("시스템", "postmaster@cesco.co.kr"); // 발신자
                    //mailInfos.ToEmails.Add(new Cesco.FW.Global.Groupware.EmailAddress("", dRow["RNMG"].ToString())); // 수신자
                    mailInfos.ToEmails.Add(new Cesco.FW.Global.Groupware.EmailAddress("", "BH0719@cesco.co.kr")); // 수신자

                    var result = Cesco.FW.Global.Groupware.MailController.Send(mailInfos);

                    // 결과값 처리
                    if (!result.Result)
                    {
                        MessageBox.Show(result.Message);
                    }
                    else
                    {
                        Dbconnection();

                        try
                        {
                            dbA.LocalInfo = new Cesco.FW.Global.DBAdapter.LocalInfo("99999", System.Reflection.MethodBase.GetCurrentMethod());
                            dbA.BindingConfig.ReceiveTimeout = new TimeSpan(0, 0, 40);
                            dbA.BindingConfig.CloseTimeout = new TimeSpan(0, 0, 40);
                            dbA.BindingConfig.TimeOut = ConfigurationDetail.TimeOuts.MINUTE10;
                            dbA.BindingConfig.SendTimeout = new TimeSpan(0, 15, 0);

                            dbA.Procedure.ProcedureName = "CESNET2.dbo.Usp_Csn_Set_DisInfection_History";

                            dbA.Procedure.ParamAdd("@GUBN", pGubn);
                            dbA.Procedure.ParamAdd("@RSEQ", dRow["RSEQ"].ToString());

                            dbA.ProcedureToDataSetCompress();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            this.Cursor = Cursors.Default;
                        }
                    }

                    mailInfos = null;
                }
            }
        }
    }
}
