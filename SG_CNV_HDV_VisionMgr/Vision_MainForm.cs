using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.IO;
using System.Xml;
using System.Threading;

using Cognex.InSight;
using Cognex.InSight.Cell;
using Cognex.InSight.Net;

using Cognex.DataMan.SDK;
using Cognex.DataMan.SDK.Discovery;
using Cognex.DataMan.SDK.Utils;

using EasyModbus;

namespace SG_CNV_HDV_VisionMgr
{
    public partial class MainForm : Form
    {
        public String Results_handheld = null;
        public String Results_Vision = null;
        
        
        //  Vision Camera
        private Cognex.InSight.CvsInSight inSight;

        //  Vision Hand-held
        private SerSystemDiscoverer serialSysDiscovery = null;  //  DMR Discovery
        private ISystemConnector sysConnector = null;
        private DataManSystem dataManSys = null;
        private ResultCollector dataManResults = null;
        private SynchronizationContext syncContext = null;  //  thread-safe
        
        //  Modbus TCP
        private ModbusServer mbServer = new ModbusServer();


        public MainForm()
        {
            Cognex.InSight.CvsInSightSoftwareDevelopmentKit.Initialize();
            
            InitializeComponent();

            setBarcodeText("123");

            //  Vision Camera
            this.inSight = new CvsInSight();
            this.inSight.ResultsChanged += new System.EventHandler(this.insight_ResultsChanged);
            this.inSight.ConnectCompleted += new CvsConnectCompletedEventHandler(this.insight_ConnectCompleted);


            syncContext = WindowsFormsSynchronizationContext.Current;
            cbbDmrList.DropDownStyle = ComboBoxStyle.DropDownList;

            //  Modbus TCP start
             
             mbServer.Listen();

        }











        #region Local_Events


        #region EventsCameraVision
        /************************************************************************/
        /*                        Camera Vision Events                          */
        /************************************************************************/
        private void insight_ConnectCompleted(object sender, CvsConnectCompletedEventArgs e)
        {
            Console.WriteLine("Connected {0}", e.ErrorMessage);

            syncContext.Post(
                new SendOrPostCallback(
                    delegate
                    {
                        
                    }), null);
            //  Addition
            this.btnCameraDisconnect.Enabled = false;
        }


        private void insight_ResultsChanged(object sender, EventArgs e)
        {
            //            Console.WriteLine("Trigger {0}", DateTime.Now.ToShortTimeString());

            //Cognex.InSight.Cell resultCells = inSight.Results.Cells["A2"];


            //  Sample
            //             CvsCellCollection insightCells = inSight.Results.Cells.GetCells(1, 1, 4, 10);   //  Get Cells from A1 to J4
            //             Console.WriteLine(insightCells.GetCell(1, 2));  //  Access matrix[1, 2]: B1

        }
        #endregion



        #region EventsHandheld
        /************************************************************************/
        /*                           Hand-held Events                           */
        /************************************************************************/
        private void OnSerialSysDiscovered(SerSystemDiscoverer.SystemInfo sysInfo)
        {
            syncContext.Post(
                new SendOrPostCallback(
                    delegate {
                        cbbDmrList.Items.Add(sysInfo);
                        cbbDmrList.SelectedIndex = cbbDmrList.FindStringExact(sysInfo.PortName);
                    }), null);
        }

        private void OnSystemConnected(object sender, EventArgs args)
        {
            syncContext.Post(
                delegate { 
                    //  Thread Connected

                    this.btnDmrConnect.Enabled = false;
                    this.btnDmrDisconnect.Enabled = true;

                }, null);
        }

        private void OnSystemDisconnected(object sender, EventArgs args)
        {
            syncContext.Post(
                delegate
                {
                    //  Thread Disconnected

                    this.btnDmrConnect.Enabled = true;
                    this.btnDmrDisconnect.Enabled = false;

                }, null);
        }

        private void OnComplexResultArrived(object sender, ResultInfo e)
        {
            syncContext.Post(
                delegate
                {
                    //  Thread Results Arrived
                    Results_handheld = !String.IsNullOrEmpty(e.ReadString) ? e.ReadString : GetReadStringFromResultXml(e.XmlResult);

                    Console.WriteLine("Results: {0}", Results_handheld);

                }, null);
        }



        #endregion



        #region EventsForm
        /************************************************************************/
        /*                              Form Events                             */
        /************************************************************************/
        
        //  On Form Load
        private void MainForm_Load(object sender, EventArgs e)
        {
            
            //  Discover Hand-held
            serialSysDiscovery = new SerSystemDiscoverer();
            serialSysDiscovery.SystemDiscovered += new SerSystemDiscoverer.SystemDiscoveredHandler(OnSerialSysDiscovered);
            serialSysDiscovery.Discover();


        }


        //  Camera Connect
        private void btnCameraConnect_Click(object sender, EventArgs e)
        {
            if (inSight.State == CvsInSightState.NotConnected)
            {
                //  Connecting to Vision
                inSight.Connect(txtCameraIP.Text, txtUsername.Text, txtPwds.Text, true, true);
            }
            
            
            //  Addition
            this.btnCameraConnect.Enabled = false;
            this.btnCameraDisconnect.Enabled = true;
            
        }

        //  Camera Disconnect
        private void btnCameraDisconnect_Click(object sender, EventArgs e)
        {
            if (inSight.State != CvsInSightState.NotConnected)
            {
                //  Disconnect from Vision
                inSight.Disconnect();
            }
            
            //  Addition
            this.btnCameraConnect.Enabled = true;
            this.btnCameraDisconnect.Enabled = false;
        }



        

        //  DataMan Connect
        private void btnDmrConnect_Click(object sender, EventArgs e)
        {
            try
            {
                var dmrComport = cbbDmrList.Items[cbbDmrList.SelectedIndex];
                if (dmrComport is SerSystemDiscoverer.SystemInfo)
                {
                    SerSystemDiscoverer.SystemInfo serSystemInfo = dmrComport as SerSystemDiscoverer.SystemInfo;
                    SerSystemConnector internalConnector = new SerSystemConnector(serSystemInfo.PortName, serSystemInfo.Baudrate);
                    sysConnector = internalConnector;
                }
                else
                {
                    Console.WriteLine("Unavailable Selection");
                }

                dataManSys = new DataManSystem(sysConnector);
                dataManSys.DefaultTimeout = 5000;   //  ---> DataManSystem.State

                dataManSys.SystemConnected += new SystemConnectedHandler(this.OnSystemConnected);
                dataManSys.SystemDisconnected += new SystemDisconnectedHandler(this.OnSystemDisconnected);


                ResultTypes requestResultTypes = ResultTypes.ReadXml;
                dataManResults = new ResultCollector(dataManSys, requestResultTypes);
                dataManResults.ComplexResultArrived += this.OnComplexResultArrived;

                dataManSys.Connect();

                try
                {
                    dataManSys.SetResultTypes(requestResultTypes);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Results type set error: {0}", ex.ToString());
                }
            }
            catch (Exception ex)
            {

                //  Unable to connect DMR
                cleanupConnection();

                Console.WriteLine("Unable to connect Hand-held: {0}", ex.ToString());
            }
        }

        //  DataMan Disconnect
        private void btnDmrDisconnect_Click(object sender, EventArgs e)
        {
            if ((dataManSys == null) || (dataManSys.State != Cognex.DataMan.SDK.ConnectionState.Connected))
            {
                return;
            }
            dataManSys.Disconnect();
            cleanupConnection();
        }

        //  Exit Form
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        #endregion

        #endregion

        #region DMR_Addition

        private string GetReadStringFromResultXml(string resultXml)
        {
            try
            {
                XmlDocument doc = new XmlDocument();

                doc.LoadXml(resultXml);

                XmlNode full_string_node = doc.SelectSingleNode("result/general/full_string");

                if (full_string_node != null)
                {
                    XmlAttribute encoding = full_string_node.Attributes["encoding"];
                    if (encoding != null && encoding.InnerText == "base64")
                    {
                        byte[] code = Convert.FromBase64String(full_string_node.InnerText);
                        return dataManSys.Encoding.GetString(code, 0, code.Length);
                    }

                    return full_string_node.InnerText;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Get DMR String Error {0}", ex.Message);
            }

            return "";
        }


        private void cleanupConnection()
        {
            if (dataManSys != null)
	        {
                dataManSys.SystemDisconnected -= this.OnSystemDisconnected;
                dataManSys.SystemConnected -= this.OnSystemConnected;
                dataManResults.ComplexResultArrived -= this.OnComplexResultArrived;
	        }

            sysConnector = null;
            dataManSys = null;
        }

        #endregion


        #region Camera Control
        private void getCameraResults()
        {
            CvsCellCollection inSightCells = inSight.Results.Cells.GetCells(1, 1, 4, 10);
            var temp = inSightCells.GetCell(1, 1);
        }

        #endregion


        #region FormControl

        private void setBarcodeText(String barString)
        {
            char[] barArr = new char[17];
            barArr = barString.ToCharArray();

            lbl_B01.Text = barArr[0].ToString();
            lbl_B02.Text = barArr[1].ToString();
            lbl_B03.Text = barArr[2].ToString();
            lbl_B04.Text = barArr[3].ToString();
            lbl_B05.Text = barArr[4].ToString();
            lbl_B06.Text = barArr[5].ToString();
            lbl_B07.Text = barArr[6].ToString();
            lbl_B08.Text = barArr[7].ToString();
            lbl_B09.Text = barArr[8].ToString();
            lbl_B10.Text = barArr[9].ToString();
            lbl_B11.Text = barArr[10].ToString();
            lbl_B12.Text = barArr[11].ToString();
            lbl_B13.Text = barArr[12].ToString();
            lbl_B14.Text = barArr[13].ToString();
            lbl_B15.Text = barArr[14].ToString();
            lbl_B16.Text = barArr[15].ToString();
            lbl_B17.Text = barArr[16].ToString();
        }

        #endregion










    }

    
}
