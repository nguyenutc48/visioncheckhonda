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

using System.Text.RegularExpressions;
using System.Xml.Serialization;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

using Cognex.InSight;
using Cognex.InSight.Cell;
using Cognex.InSight.Net;
using Cognex.InSight.Sensor;

using Cognex.DataMan.SDK;
using Cognex.DataMan.SDK.Discovery;
using Cognex.DataMan.SDK.Utils;

using EasyModbus;

using System.Diagnostics;

namespace SaGiangVisionManager
{
    
    
    public partial class MainForm : Form
    {
        Stopwatch stopWatch01 = new Stopwatch();

        public String LogFilePath = "C:\\Users\\Public\\Documents";
        public String LogFileName = "yyyy.MM.dd";   //  "dd.MM.yyyy"

        public String configFilePath = "C:\\Users\\Public\\Documents";
        public String configFileName = "Config.xml";

        public String Results_HandheldNow = "";
        public String Results_Vision = "";
        public bool Results_OK = true;

        Int16 Fixed_barcodeLen = 18;    //  Old = 17 @Jan 13 2017
        Int16 Fixed_visionLen = 18;     //  Old = 17 @Jan 13 2017

        //  Alert variables
        bool isHandheldRead = false;    //  true when hand-held data is captured
        bool isVisionRead = false;      //  true when Vision data is captured
        
        bool VMSystemState = false; //  Enable capture data from hand-held & vision system after START VisionManager.

        
        public String[] Results_HandheldQueue = new String[2];  //  Queue hand-held data string when hand-held return more than 1 data before camera trigger.
        String displayHandheldString = "";
        int queuePosition = 0;                                  //  Return Queue position.
        bool queueLock = true;                                  //  Lock queue position
        bool isQueueUp = false;                                 //  return status of QueueUp?

        bool inCorruptSerialConnection = false;

        private int plcConnectRetry = 0;




        enum PortState
        {
            Connecting_ = 0,
            Connected_ = 1,
            Disconnecting_ = 2,
            Disconnected_ = 3,
            LostConnection_ = 4,
            ReConnecting_ = 5,
        }

        enum SensorState
        {
            Connecting_ = 0,
            Connected_ = 1,
            Disconnecting_ = 2,
            Disconnected_ = 3,
            NotConnect_ = 4,
            Offline_ = 5,
            Online_= 6
        }

        
        

        #region Variable Vision
        static String[] SensorState_t = new String[] { "Connecting", "Connected", "Disconnecting", "Disconnected", "Not Connected", "Offline", "Online" };
        int VisionConnectionState = (int)SensorState.Disconnected_;
        #endregion

        #region Variable Serial Port
        /************************************************************************/
        /*                           Serial Variables                           */
        /************************************************************************/
        String RxString = "";
        delegate void SetTextCallBack(string text);
        String[] baudArray = new String[] { "1200", "2400", "4800", "9600", "19200", "38400", "57600", "115200", "230400" };

        //  Connection State
        static String[] SerialPortState_t = new String[] { "Connecting", "Connected", "Disconnecting", "Disconnected", "Lost Connection", "Re-connecting" };
        

        int SerialPortState = (int)PortState.Disconnected_;

        #endregion

        #region Variable Saved
        /************************************************************************/
        /*                           Saved Variables                            */
        /************************************************************************/
        Infomation formInfo;
        #endregion

        #region Variable Shift

        /************************************************************************/
        /*                           Shift Variables                            */
        /************************************************************************/
        int Ca01_Hour = 6;
        int Ca01_Min = 0;

        int Ca02_Hour = 15;
        int Ca02_Min = 0;
        #endregion

        #region Variable Spreadsheet
        /************************************************************************/
        /*                        Spreadsheet Variables                         */
        /************************************************************************/
        enum ColPosition
        {
            Col_No = 0,
            Col_Date = 1,
            Col_Time = 2,
            Col_Shift = 3,
            Col_Model = 4,
            Col_BarcodeData = 5,
            Col_VisionData = 6,
            Col_Result = 7
        }

        static String[] tableCols = new String[] { "No.", "Date", "Time", "Shift", "Model", "Barcode", "Vision", "OK/NG" };
        String dateFormat = "dd-MMM-yyyy";
        String timeFormat = "hh:mm:ss";

        SLDocument SSL = new SLDocument();
        DataTable resultTable = new DataTable("resultsTable");
        #endregion

        #region Variable Camera
        /************************************************************************/
        /*                          Camera Variables                            */
        /************************************************************************/
        //  Vision Camera
        private Cognex.InSight.CvsInSight inSight;

        private bool fullAccess = false;

        bool[] checkList = new bool[5];
        bool[] ResultCharHeigh = new bool[17];
        bool[] ResultCharWidth = new bool[17];
        bool[] ResultCharAngle = new bool[17];
        bool[] ResultCharDistance = new bool[16];
        bool ResultCharLineAngle = false;

        private List<String> jobnames = new List<string> { };
        private List<String> jobnamesincam = new List<string> { };
        private List<String> modelnames = new List<string> { };
        private string[] mFileList = { };
        private int loaddone = 0;
        private string plc_address_start = "M500";
        private int index_select_before = -1;
        private CvsCell resultCell = null;
        private CvsCellCollection inSightCells = null;
        private CvsCellCollection inSightResultCells = null;
        private CvsCellCollection inSightCellCheckParameter = null;

        #endregion

        #region Variable Hand-held
        /************************************************************************/
        /*                        Hand-held Variables                           */
        /************************************************************************/
        //  Vision Hand-held
        private SerSystemDiscoverer serialSysDiscovery = null;//  DMR Serial Discovery
        private EthSystemDiscoverer ethSystemDiscoverer = null;    //  DMR Ethernet Discovery
        private ISystemConnector sysConnector = null;
        private DataManSystem dataManSys = null;
        private ResultCollector dataManResults = null;
        private SynchronizationContext syncContext = null;  //  thread-safe
        private DmccResponse DMR_ResponseResults;
        #endregion

        #region Variable Modbus
        //  Modbus TCP
        private ModbusServer mbServer = new ModbusServer();
        #endregion


        public MainForm()
        {
            //Cognex.InSight.CvsInSightSoftwareDevelopmentKit.Initialize();
            
            InitializeComponent();

            
            
            

            //  Maximize Window
            this.WindowState = FormWindowState.Maximized;
            
            //  Init SpreadSheet View
            initTable();
            clearResultLabel(true);
            clearResultOK_NG();
            displayResultReset();

            //  Disable Add Model
            this.grpAddModel.Enabled = false;
            //  Disable SystemStop
            this.btnStopSystem.Enabled = false;

            //  Vision Camera
            this.inSight = new CvsInSight();
            this.inSight.ResultsChanged += new System.EventHandler(this.insight_ResultsChanged);
            this.inSight.ConnectCompleted += new CvsConnectCompletedEventHandler(this.insight_ConnectCompleted);
            this.inSight.StateChanged += new CvsStateChangedEventHandler(this.insight_StateChanged);

            //  Synchronize GUI Objects with a-synchronize thread.
            syncContext = WindowsFormsSynchronizationContext.Current;


            //  Start ModebusTCP Server
            mbServer.Listen();


            //  Clear all display
            clearResultLabel(true);
            clearResultOK_NG();
            displayResultReset();


            //  Vision State
            VisionConnectionState = (int)SensorState.Disconnected_;
            displayVisionConnectionState();

            //  Hand-held State
            SerialPortState = (int)PortState.Disconnected_;
            displayDMRportState();

            setPlcTextStatus("Disconnected", System.Drawing.Color.Red);


            displayFileVersion();
            
        }


        #region FormEvents

        //  Form Load Event
        private void MainForm_Load(object sender, EventArgs e)
        {
            //  Discover Hand-held
            HandheldDiscover();
            PlcDiscovery();
            
            //  using Info
            formInfo = new Infomation();
            loadConfig();
            loadModel();
            txtJobInput.ReadOnly = true;
            if (modelnames.Count != 0)
                lblAngle.Text = jobnames[0];
            //this.btn_Confirm.BackColor = System.Drawing.Color.Yellow;
            //this.btn_Confirm.Text = "Confirm";
            //  Auto-Connect
            //             if (this.checkAutoConnect.CheckState == CheckState.Checked)
            //             {
            //                 syncContext.Post(
            //                 new SendOrPostCallback(
            //                     delegate
            //                     {
            //                         this.btnCamConnect_Click(null, null);
            //                         this.btnDmrConnect_Click(null, null);
            //                     }), null);
            //             }

            toggleSwitch_ShiftAutoMan_Toggled(sender, e);

            //  Re-load Excel files.
            loadSpreadsheet();
        }


        private void tmrStartButton_Tick(object sender, EventArgs e)
        {
            this.tmrStartButton.Enabled = false;
            //  Start Automatically
            this.btnStartSystem_Click(null, null);
        }

        //  Form Closing
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Bạn có chắc muốn đóng chương trình không?", "Cảnh báo đóng chương trình", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {

                this.formInfo.CameraIP = txtCamIP.Text;
                this.formInfo.CameraUsr = txtCamUsr.Text;
                this.formInfo.CameraPwds = txtCamPwds.Text;
                this.formInfo.CurrentModel = index_select_before.ToString();

                //this.inSight.ResultsChanged -= new System.EventHandler(this.insight_ResultsChanged);
                //this.inSight.ConnectCompleted -= new CvsConnectCompletedEventHandler(this.insight_ConnectCompleted);
                //this.inSight.StateChanged -= new CvsStateChangedEventHandler(this.insight_StateChanged);

                //  Save config to XML
                saveConfig();

                //  Clear DMR object buffer
                cleanupConnection();
            }
            else
            {
                e.Cancel = true;
                return;
            }
            
            
        }

        //  Connect to DMR
        private void btnDmrConnect_Click(object sender, EventArgs e)
        {
            ConnectDMR();
        }

        private void ConnectDMR()
        {
            try
            {
                var dmrComport = cbbDmrList.Items[cbbDmrList.SelectedIndex];

                if (dmrComport is SerSystemDiscoverer.SystemInfo)
                {
                    SerSystemDiscoverer.SystemInfo serSystemInfo = dmrComport as SerSystemDiscoverer.SystemInfo;
                    SerSystemConnector internalConnector = new SerSystemConnector(serSystemInfo.PortName, serSystemInfo.Baudrate);
                    sysConnector = internalConnector;

                    //  Normal Serial comport Connector
                    serialPort.PortName = serSystemInfo.PortName;
                    serialPort.BaudRate = Convert.ToInt32(cbbDmrBaud.SelectedItem);
                    serialPort.DtrEnable = true;
                }
                else if (dmrComport is EthSystemDiscoverer.SystemInfo)
                {
                    EthSystemDiscoverer.SystemInfo ethSystemInfo = dmrComport as EthSystemDiscoverer.SystemInfo;
                    EthSystemConnector internalConnector = new EthSystemConnector(ethSystemInfo.IPAddress, ethSystemInfo.Port);

                    internalConnector.UserName = "admin";
                    internalConnector.Password = "";

                    sysConnector = internalConnector;
                }



                if (chkbox_IsDMR.CheckState == CheckState.Checked)  //  DataMan devices on the System
                {
                    dataManSys = new DataManSystem(sysConnector);
                    dataManSys.DefaultTimeout = 1000;   //  ---> DataManSystem.State

                    dataManSys.SystemConnected += new SystemConnectedHandler(this.OnSystemConnected);
                    dataManSys.SystemDisconnected += new SystemDisconnectedHandler(this.OnSystemDisconnected);
                    dataManSys.ReadStringArrived += new ReadStringArrivedHandler(this.OnReadStringArrived);
                    dataManSys.StatusEventArrived += new StatusEventArrivedHandler(this.OnStatusEventArrived);
                    dataManSys.BinaryDataTransferProgress += new BinaryDataTransferProgressHandler(this.OnBinaryDataTransferProgress);

                    //ResultTypes requestResultTypes = ResultTypes.ReadXml;
                    ResultTypes requestResultTypes = ResultTypes.ReadString;
                    dataManResults = new ResultCollector(dataManSys, requestResultTypes);
                    dataManResults.ComplexResultArrived += this.OnComplexResultArrived;

                    dataManSys.Connect();

                    //  DataMan in connecting state.
                    this.SerialPortState = (int)PortState.Connecting_;
                    this.displayDMRportState();

                    try
                    {
                        dataManSys.SetResultTypes(requestResultTypes);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Results type set error: {0}", ex.ToString());
                    }

                    
                }
                else if (chkbox_IsDMR.CheckState == CheckState.Unchecked)   //  Other devices
                {
                    if (!serialPort.IsOpen)
                    {
                        Console.WriteLine("COM {0} | {1}", serialPort.PortName.ToString(), serialPort.BaudRate.ToString());
                        serialPort.Open();
                    }

                    if (serialPort.IsOpen)
                    {
                        //  Display DataMan Connected.
                        this.SerialPortState = (int)PortState.Connected_;
                        this.displayDMRportState();
                    }
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine("Barcode Connect Ex: {0}", ex.Message);
            }
        }

        private void DisconnectDMR()
        {
            if (chkbox_IsDMR.CheckState == CheckState.Checked)
            {
//                 if ((dataManSys == null) || (dataManSys.State != Cognex.DataMan.SDK.ConnectionState.Connected))
//                 {
//                     return;
//                 }

                dataManSys.Disconnect();
                cleanupConnection();

                Console.WriteLine("[{0}] Disconnected Serial Port", DateTime.Now.ToString("H:mm:ss.fff"));

                //  Serial Port is in disconnected state
                this.SerialPortState = (int)PortState.Disconnected_;
                this.displayDMRportState();
            }
            else if (chkbox_IsDMR.CheckState == CheckState.Unchecked)
            {
                try
                {
                    if (serialPort.IsOpen)
                    {
                        serialPort.DiscardInBuffer();
                        serialPort.Close();
                        serialPort.Dispose();

                        this.SerialPortState = (int)PortState.Disconnected_;
                        this.displayDMRportState();
                    }
                }
                catch
                {
                    MessageBox.Show("Cannot close this COM port!");
                }
            }
        }

        //  Disconnect from DMR
        private void btnDmrDisconnect_Click(object sender, EventArgs e)
        {
            DisconnectDMR();
        }

        //  Connect Vision Camera
        private void btnCamConnect_Click(object sender, EventArgs e)
        {
            
            if (inSight.State == CvsInSightState.NotConnected)
            {
                //  Connecting to Vision
                try
                {
                    inSight.Connect(this.txtCamIP.Text, this.txtCamUsr.Text, this.txtCamPwds.Text, true, true);
                                      
                        //  Display Connection state to GUI
                    VisionConnectionState = (int)SensorState.Connecting_;
                    displayVisionConnectionState();
                }
                catch (Exception ex)
                {
                    //  Display Connection state to GUI
                    VisionConnectionState = (int)SensorState.NotConnect_;
                    displayVisionConnectionState();
                    // MessageBox.Show(ex.Message);
                }
                
            }

            
        }

        //  Disconnect Vision Camera
        private void btnCamDisconnect_Click(object sender, EventArgs e)
        {
            if (inSight.State != CvsInSightState.NotConnected)
            {
                //  Disconnect from Vision
                inSight.Disconnect();
                fullAccess = false;
                //this.grpAddModel.Enabled = true;

                //  Set Vision State
                VisionConnectionState = (int)SensorState.Disconnected_;
            }

            //  Display Vision Not Connect or Disconnect
            //display_VisionNotConnect();
            displayVisionConnectionState();
        }

        //  Is DataMan Devices on the system.
        private void chkbox_IsDMR_CheckedChanged(object sender, EventArgs e)
        {
            HandheldDiscover();
            if (chkbox_IsDMR.CheckState == CheckState.Checked)
            {
                cbbDmrBaud.Enabled = false;
                formInfo.IsDMRCheckState = true;
            }
            else if (chkbox_IsDMR.CheckState == CheckState.Unchecked)
            {
                cbbDmrBaud.Enabled = true;
                formInfo.IsDMRCheckState = false;
            }

            
        }

        //  Auto Connect
        private void checkAutoConnect_CheckedChanged(object sender, EventArgs e)
        {
            if (checkAutoConnect.CheckState == CheckState.Checked)
            {
                formInfo.autoLogin = true;
                //  Start Automatically
                if (this.tmrStartButton.Enabled == false)
                {
                    this.tmrStartButton.Enabled = true;
                }
            }
            else if(checkAutoConnect.CheckState == CheckState.Unchecked)
            {
                formInfo.autoLogin = false;
                this.tmrStartButton.Enabled = false;    
            }
        }

        //  START System
        private void btnStartSystem_Click(object sender, EventArgs e)
        {
            if (modelnames.Count == 0)
            {
                Console.WriteLine("Khong co model");
                return;
            }
            this.btnStartSystem.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            this.btnStopSystem.Enabled = true;
            this.btnStartSystem.Enabled = false;
            this.grpAddModel.Enabled = false;
            
            SystemStart();
        }

        //  STOP System
        private void btnStopSystem_Click(object sender, EventArgs e)
        {
            this.btnStartSystem.Enabled = true;
            this.btnStopSystem.Enabled = false;
            this.grpAddModel.Enabled = true;
            SystemStop();
        }
        #endregion

        #region Events Camera
        /************************************************************************/
        /*                        Camera Vision Events                          */
        /************************************************************************/

        //  Camera State Changed Events
        private void insight_StateChanged(object sender, CvsStateChangedEventArgs e)
        {
            syncContext.Post(
                new SendOrPostCallback(
                    delegate
                    {
                        switch (inSight.State)
                        {
                            case CvsInSightState.NotConnected:
                                stt_CameraStatus.Caption = "Not Connected";
                                break;
                            case CvsInSightState.Offline:
                                inSight.SoftOnline = true;
                                stt_CameraStatus.Caption = "Offline";
                                break;
                            case CvsInSightState.Online:
                                stt_CameraStatus.Caption = "Online";
                                break;
                            default:
                                break;
                        }
                    }), null);
        }

        //  Camera Connect Complete
        private void insight_ConnectCompleted(object sender, CvsConnectCompletedEventArgs e)
        {
            Console.WriteLine("Connected {0}, Username: {1}, State: {2}", e.ErrorMessage, inSight.Username, inSight.State.ToString());

            if (String.IsNullOrEmpty(e.ErrorMessage))
            {
                this.VisionConnectionState = (int)SensorState.Connected_;
                LoadJogFileFromInSight();
                if (inSight.Username == "admin")
                {
                    fullAccess = true;
                }
                else
                {
                    fullAccess = false;
                }
            }
            else
            {
                this.VisionConnectionState = (int)SensorState.NotConnect_;
            }



            


            syncContext.Post(
                new SendOrPostCallback(
                    delegate
                    {
                        displayVisionConnectionState();
                    }), null);


        }

        //  Camera Results Change, trig every acquisition completed
        private void insight_ResultsChanged(object sender, EventArgs e)
        {
            
            
            displayResultReset();
            clearResultLabel(false);
            
            //  Get a data table of cells
            //  Get table range from B1 to T6
            //  1: char results
            //  2: char height
            //  3: char width
            //  4: angle of single char
            //  5: Distance from center to center
            //  6: Angle of whole line string.


            Int16 positionCharResults = 1;
            Int16 positionCharHeight = 2;
            Int16 positionCharWidth = 3;
            Int16 positionCharAngle = 4;
            Int16 positionCenter2Center = 5;
            Int16 positionLineStrAngle = 6;

            /*
            //  Chassis Parameters
            CvsCellCollection inSightCells = inSight.Results.Cells.GetCells(1, 1, 6, 19);   //  B1 [1, 1] to T6 [6, 19] [row, col]
            CvsCellCollection inSightResultCells = inSight.Results.Cells.GetCells(202, 2, 206, 18);
            CvsCellCollection inSightCellCheckParameter = inSight.Results.Cells.GetCells(181, 2, 181, 6);
            */

            //  Engine Block Parameters
            inSightCells = inSight.Results.Cells.GetCells(1, 1, 6, 19);   //  B1 [1, 1] to T6 [6, 19] [row, col]
            inSightResultCells = inSight.Results.Cells.GetCells(202, 2, 206, 18);
            inSightCellCheckParameter = inSight.Results.Cells.GetCells(181, 2, 181, 6);

            if (inSightResultCells == null)
                Console.WriteLine("inSightResultCells is null");
            if (inSightCells == null)
                Console.WriteLine("inSightCells is null");
            if (inSightCellCheckParameter == null)
                Console.WriteLine("inSightCellCheckParameter is null");

            resultCell = inSight.Results.Cells["A183"];
            //sua ngay 9/7/2018 update loi khong lay duoc du lieu
            if (resultCell != null)
            {
                if (Convert.ToDouble(resultCell.ToString()) == 1)
                {
                    Results_OK = true;
                }
                else
                {
                    Results_OK = false;
                }


                //  Get Results String from Vision System.
                String visionReadResults = String.Empty;
                //  Get Results of Char High from vision system
                short tmpVisionLens = Fixed_visionLen;
                Fixed_visionLen = 13;// new character length

                double[] charHeight = new double[Fixed_visionLen];
                double[] charWidth = new double[Fixed_visionLen];
                double[] charAngle = new double[Fixed_visionLen];
                double[] charCenter2Center = new double[Fixed_visionLen - 1];
                double lineStrAngle;




                //  Character Limit
                double UpLimit_charHeight = 5.2;
                double LowLimit_charHeight = 4.8;

                double UpLimit_charWidth = 3.2;
                double LowLimit_charWidth = 2.8;

                double UpLimit_charAngle = 15.0;
                double LowLimit_charAngle = 0.0;

                double UpLimit_DistanceC2C = 7.0;
                double LowLimit_DistanceC2C = 4.0;

                double UpLimit_LineStrAngle = 15.0;
                double LowLimit_LineStrAngle = 0.0;


                UpLimit_charHeight = Convert.ToDouble(inSightCells.GetCell(2, 18).ToString());
                LowLimit_charHeight = Convert.ToDouble(inSightCells.GetCell(2, 19).ToString());
                UpLimit_charWidth = Convert.ToDouble(inSightCells.GetCell(3, 18).ToString());
                LowLimit_charWidth = Convert.ToDouble(inSightCells.GetCell(3, 19).ToString());
                UpLimit_charAngle = Convert.ToDouble(inSightCells.GetCell(4, 18).ToString());
                LowLimit_charAngle = Convert.ToDouble(inSightCells.GetCell(4, 19).ToString());
                UpLimit_DistanceC2C = Convert.ToDouble(inSightCells.GetCell(5, 18).ToString());
                LowLimit_DistanceC2C = Convert.ToDouble(inSightCells.GetCell(5, 19).ToString());
                UpLimit_LineStrAngle = Convert.ToDouble(inSightCells.GetCell(6, 18).ToString());
                LowLimit_LineStrAngle = Convert.ToDouble(inSightCells.GetCell(6, 19).ToString());


                for (int i = 1; i <= Fixed_visionLen; i++)
                {
                    /************************************************************************/
                    //  Results Vision String
                    if (inSightCells.GetCell(positionCharResults, i).Error.ToString() == "True")
                    {
                        //  if cell is #ERR then input "#"
                        visionReadResults += "#";
                    }
                    else
                    {
                        //  if cell is OK then input cell result string
                        visionReadResults += inSightCells.GetCell(positionCharResults, i);
                    }


                    /************************************************************************/
                    //  Results Vision Height
                    if (inSightCells.GetCell(positionCharHeight, i).Error.ToString() == "True")
                    {
                        //  if cell is #ERR then input -1;
                        charHeight[i - 1] = -1;
                    }
                    else
                    {
                        //  if cell is OK then input results
                        charHeight[i - 1] = Convert.ToDouble(inSightCells.GetCell(positionCharHeight, i).ToString());

                        //Console.WriteLine("Char Height Sensor {0} | Internal {1}", inSightCells.GetCell(2, i), charHeight[i - 1]);
                    }


                    /************************************************************************/
                    //  Results Vision Width
                    if (inSightCells.GetCell(positionCharWidth, i).Error.ToString() == "True")
                    {
                        //  if cell is #ERR then input -1;
                        charWidth[i - 1] = -1;
                    }
                    else
                    {
                        //  if cell is OK then input results
                        charWidth[i - 1] = Convert.ToDouble(inSightCells.GetCell(positionCharWidth, i).ToString());
                    }

                    /************************************************************************/
                    //  Results Vision Angle
                    if (inSightCells.GetCell(positionCharAngle, i).Error.ToString() == "True")
                    {
                        //  if cell is #ERR then input -1;
                        charAngle[i - 1] = 90;
                    }
                    else
                    {
                        //  if cell is OK then input results
                        charAngle[i - 1] = Convert.ToDouble(inSightCells.GetCell(positionCharAngle, i).ToString());
                    }

                    /************************************************************************/
                    //  Results Vision Distance from Center to Center
                    if (i < Fixed_visionLen)
                    {
                        if (inSightCells.GetCell(positionCenter2Center, i).Error.ToString() == "True")
                        {
                            //  if cell is #ERR then input -1;
                            charCenter2Center[i - 1] = -1;
                        }
                        else
                        {
                            //  if cell is OK then input results
                            charCenter2Center[i - 1] = Convert.ToDouble(inSightCells.GetCell(positionCenter2Center, i).ToString());

                            //Console.WriteLine("Center2Center {0} | {1}", inSightCells.GetCell(positionCenter2Center, i), charCenter2Center[i - 1]);
                        }
                    }


                    /************************************************************************/
                    /*                          Results OK/NG Height                        */
                    if (inSightResultCells.GetCell(202, i + 1).Error.ToString() == "True")
                    {
                        //  if cell is #ERR then input -1;
                        ResultCharHeigh[i - 1] = false;
                    }
                    else
                    {
                        //  if cell is OK then input results
                        if (Convert.ToDouble(inSightResultCells.GetCell(202, i + 1).ToString()) == 1.0)
                        {
                            ResultCharHeigh[i - 1] = true;
                        }
                        else
                        {
                            ResultCharHeigh[i - 1] = false;
                        }
                    }

                    /************************************************************************/
                    /*                          Results OK/NG Width                         */
                    if (inSightResultCells.GetCell(203, i + 1).Error.ToString() == "True")
                    {
                        //  if cell is #ERR then input -1;
                        ResultCharWidth[i - 1] = false;
                    }
                    else
                    {
                        //  if cell is OK then input results
                        if (Convert.ToDouble(inSightResultCells.GetCell(203, i + 1).ToString()) == 1.0)
                        {
                            ResultCharWidth[i - 1] = true;
                        }
                        else
                        {
                            ResultCharWidth[i - 1] = false;
                        }
                    }

                    /************************************************************************/
                    /*                          Results OK/NG Angle                         */
                    if (inSightResultCells.GetCell(204, i + 1).Error.ToString() == "True")
                    {
                        //  if cell is #ERR then input -1;
                        ResultCharAngle[i - 1] = false;
                    }
                    else
                    {
                        //  if cell is OK then input results
                        if (Convert.ToDouble(inSightResultCells.GetCell(204, i + 1).ToString()) == 1.0)
                        {
                            ResultCharAngle[i - 1] = true;
                        }
                        else
                        {
                            ResultCharAngle[i - 1] = false;
                        }
                    }

                    /************************************************************************/
                    /*                          Results OK/NG Distance                      */
                    if (i < Fixed_visionLen)
                    {
                        if (inSightResultCells.GetCell(205, i + 1).Error.ToString() == "True")
                        {
                            //  if cell is #ERR then input -1;
                            ResultCharDistance[i - 1] = false;
                        }
                        else
                        {
                            //  if cell is OK then input results
                            if (Convert.ToDouble(inSightResultCells.GetCell(205, i + 1).ToString()) == 1.0)
                            {
                                ResultCharDistance[i - 1] = true;
                            }
                            else
                            {
                                ResultCharDistance[i - 1] = false;
                            }
                        }
                    }



                }

                Fixed_visionLen = tmpVisionLens;    //  Return to default

                /************************************************************************/
                //  Results Vision String Line Angle
                if (inSightCells.GetCell(positionCharAngle, 1).Error.ToString() == "True")
                {
                    //  if cell is #ERR then input -1;
                    lineStrAngle = 90;
                }
                else
                {
                    //  if cell is OK then input results
                    lineStrAngle = Convert.ToDouble(inSightCells.GetCell(positionLineStrAngle, 1).ToString());
                }


                /************************************************************************/
                /*                        Results OK/NG Line Angle                      */
                if (inSightResultCells.GetCell(206, 2).Error.ToString() == "True")
                {
                    //  if cell is #ERR then input -1;
                    ResultCharLineAngle = false;
                }
                else
                {
                    //  if cell is OK then input results
                    if (Convert.ToDouble(inSightResultCells.GetCell(206, 2).ToString()) == 1.0)
                    {
                        ResultCharLineAngle = true;
                    }
                    else
                    {
                        ResultCharLineAngle = false;
                    }
                }



                /************************************************************************/
                //  Compare the new with old one ---> if they are same ---> nothing change ---> no update requires!
                //  Else, it will do below.
                if (String.Compare(Results_Vision, visionReadResults) != 0)
                {
                    Results_Vision = visionReadResults;
                }

                /************************************************************************/
                //  Add results to data table



                Console.WriteLine("Results [{0}]: {1}", visionReadResults.Length, visionReadResults);




                syncContext.Post(
                    new SendOrPostCallback(
                        delegate
                        {
                            Console.WriteLine("[{0}] IsHandheldRead [{1}] | IsQueueUp [{2}]", DateTime.Now.ToString("H:mm:ss.fff"), isHandheldRead, isQueueUp);

                            if (isHandheldRead)
                            {
                                isHandheldRead = false;         //  Disable hand-held read indicator
                            isVisionRead = true;    //  Vision system has been triggered

                            addVisionData(Results_Vision);  //  add result vision into Spead-sheet
                            addResult_OK_NG(Results_OK);    //  add OK/NG into Spreadsheet


                            if (isQueueUp)
                                {
                                    TMR_DisplayDelay.Enabled = true;    //  Enable delay Display Timer
                            }
                                else
                                {
                                    TMR_DisplayDelay.Enabled = false;    //  Enable delay Display Timer
                            }



                            //  Show check/un-check Inspection List.
                            displayCheckParameter(inSightCellCheckParameter);


                            //  Print Vision data to Screen
                            setVisionText_v2(Results_Vision);
                            //  DISPLAY string compare between vision & bar-code data
                            displayCompareResult();

                            //  Print OK/NG Result
                            if (Results_OK)
                                {
                                    displayResultOK();
                                }
                                else
                                {
                                    displayResultNG();
                                }

                            //  Print Line Angle to Screen
                            //inputLineStrAngle(lineStrAngle, LowLimit_LineStrAngle, UpLimit_LineStrAngle);
                            inputLineStrAngle(lineStrAngle, ResultCharLineAngle);

                                short tmpFixedVisionLens = 13;

                            //for (int i = 1; i <= Fixed_visionLen; i++)
                            for (int i = 1; i <= tmpFixedVisionLens; i++)
                                {
                                //inputCharSpecs(i, charHeight[i - 1], charWidth[i - 1], charAngle[i - 1], UpLimit_charHeight, LowLimit_charHeight, UpLimit_charWidth, LowLimit_charWidth, UpLimit_charAngle, LowLimit_charAngle);
                                inputCharSpecs(i, charHeight[i - 1], charWidth[i - 1], charAngle[i - 1], ResultCharHeigh[i - 1], ResultCharWidth[i - 1], ResultCharAngle[i - 1]);
                                    if (i < /*Fixed_visionLen*/ tmpFixedVisionLens)
                                    {
                                    //inputDistanceCenter2Center(i, charCenter2Center[i - 1], UpLimit_DistanceC2C, LowLimit_DistanceC2C);
                                    inputDistanceCenter2Center(i, charCenter2Center[i - 1], ResultCharDistance[i - 1]);
                                    }
                                }


                            //  Automatic save Spreadsheet file into specified location.
                            saveSpreadsheet();

                            }


                        }), null);

            }
            else
                Console.WriteLine("resultCell is null");

        }
        #endregion

        #region Events PLC
        private void PlcDiscovery()
        {
            string[] portName = System.IO.Ports.SerialPort.GetPortNames();
            this.cbbPlcPortList.Items.Clear();
            for (int i = 0; i < portName.Length; i++)
            {
                if (!cbbPlcPortList.Items.Contains(portName[i])) //  Prevent add duplicate item
                {
                    cbbPlcPortList.Items.Add(portName[i]);
                    cbbPlcPortList.SelectedIndex = 0;
                }
            }



        }


        private void btnPlcDisconnect_Click(object sender, EventArgs e)
        {
            try
            {
                int iReturnCode = this.axActProgType1.Close();
                if (iReturnCode != 0)
                {
                    //  Error to Connect to PLC
                    Console.WriteLine("Open Err: 0x{0:X}", iReturnCode);

                    syncContext.Post(
                        delegate
                        {
                            setPlcTextStatus("#ERR 0x" + iReturnCode.ToString("X8"), System.Drawing.Color.Red);
                        }, null);
                }
                else
                {
                    this.btnPlcConnect.Enabled = true;
                    this.btnPlcDisconnect.Enabled = false;
                    this.cbbPlcPortList.Enabled = true;

                    setPlcTextStatus("Disconnected", System.Drawing.Color.Red);

                    if (VMSystemState)
                    {
                        tmrPlcUpdate.Enabled = false;
                    }
                }
                
            }
            catch (Exception ex)
            {
                
                
            }
            
        }

        private void btnPlcConnect_Click(object sender, EventArgs e)
        {
            plcPerformConnect();
        }

        private void plcPerformConnect()
        {
            
            string portName = this.cbbPlcPortList.SelectedItem.ToString();

            formInfo.PlcComport = portName;

            portName = portName.Replace("COM", "");
            this.axActProgType1.ActPortNumber = Convert.ToInt16(portName);    //  Port Number
            this.axActProgType1.ActProtocolType = 0x04;   //  Serial Communication
            this.axActProgType1.ActUnitType = 0x0F;       //  FXCPU RS422 port Direct connection
            this.axActProgType1.ActCpuType = 520;         //  FX3U(C)
            //this.axActFXCPU1.ActPortNumber = Convert.ToInt16(portName);
            //this.axActFXCPU1.ActCpuType = 520; //   FX3U(C)

            //Console.WriteLine("PLC Port: {0}", axActProgType1.ActPortNumber);

            try
            {
                int iReturnCode = this.axActProgType1.Open();
                if (iReturnCode != 0)
                {
                    //  Error to Connect to PLC
                    // Console.WriteLine("[{0}] Open Err: 0x{1:X}", DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss.fff"),iReturnCode);

                    syncContext.Post(
                        delegate
                        {
                            setPlcTextStatus("#ERR 0x" + iReturnCode.ToString("X8"), System.Drawing.Color.Red);
                            plcConnectRetry++;
                            if (plcConnectRetry < 400)
                            {
                                plcPerformConnect();
                                //Console.WriteLine("[{0}] Retry {1} (times)", DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss.fff"), plcConnectRetry);
                            }
                        }, null);

                }
                else
                {
                    plcConnectRetry = 0;
                    
                    this.btnPlcConnect.Enabled = false;
                    this.btnPlcDisconnect.Enabled = true;
                    this.cbbPlcPortList.Enabled = false;

                    setPlcTextStatus("Connected", System.Drawing.Color.Lime);

                    if (VMSystemState)
                    {
                        tmrPlcUpdate.Enabled = true;
                    }
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
        }

        private void tmrPlcUpdate_Tick(object sender, EventArgs e)
        {
            if (VMSystemState)
            {
                if (this.toggleSwitch_ShiftAutoMan.IsOn)
                {
                    var plc_start = 500;
                    for (int i = 0; i < modelnames.Count; i++)
                    {
                        
                        if(PLCReadDevice("M"+(plc_start+i).ToString())>0)
                        {
                            //Console.WriteLine("M" + (plc_start + i).ToString());
                            if (cbb_ModelList.SelectedIndex != i)
                            {
                                this.cbb_ModelList.SelectedIndex = i;
                                if (index_select_before != i)
                                {
                                    this.btn_Confirm.BackColor = System.Drawing.Color.Red;
                                    this.btn_Confirm.Text = "Confirm";
                                }
                                else
                                {
                                    this.btn_Confirm.BackColor = System.Drawing.Color.Lime;
                                    this.btn_Confirm.Text = "Confirmed";
                                }
                            }

                            
                        }
                    }
                    if(PLCReadDevice("M1000")>0)
                    {
                        //btnStopSystem_Click(null, null);
                        if (cbb_ModelList.SelectedIndex != index_select_before)
                        {
                            index_select_before = cbb_ModelList.SelectedIndex;
                            if(jobnames.Count != 0)
                                OpenJobFileToInSight(jobnames[cbb_ModelList.SelectedIndex]);
                            this.btn_Confirm.BackColor = System.Drawing.Color.Lime;
                            this.btn_Confirm.Text = "Confirmed";
                        }
                        //btnStartSystem_Click(null, null);
                    }


                    /*if (readPlcModel_M10() != 0)
                    {
                        if (this.cbb_ModelList.Items.Count > 0)
                        {
                            this.cbb_ModelList.SelectedIndex = 0;
                        }
                    }

                    if (readPlcModel_M11() != 0)
                    {
                        if (this.cbb_ModelList.Items.Count > 1)
                        {
                            this.cbb_ModelList.SelectedIndex = 1;
                        }
                    }

                    if (readPlcModel_M12() != 0)
                    {
                        if (this.cbb_ModelList.Items.Count > 2)
                        {
                            this.cbb_ModelList.SelectedIndex = 2;
                        }
                    }

                    if (readPlcModel_M13() != 0)
                    {
                        if (this.cbb_ModelList.Items.Count > 3)
                        {
                            this.cbb_ModelList.SelectedIndex = 3;
                        }
                    }

                    if (readPlcModel_M14() != 0)
                    {
                        if (this.cbb_ModelList.Items.Count > 4)
                        {
                            this.cbb_ModelList.SelectedIndex = 4;
                        }
                    }

                    if (readPlcModel_M15() != 0)
                    {
                        if (this.cbb_ModelList.Items.Count > 5)
                        {
                            this.cbb_ModelList.SelectedIndex = 5;
                        }
                    }*/
                }
            }
        }

        int readPlcModel_M10()
        {
            int resultDeviceValue = 0;

            try
            {
                int iReturnCode = this.axActProgType1.GetDevice("M10", out resultDeviceValue);
                if (iReturnCode != 0)
                {
                    //  Error
                    Console.WriteLine("Error while reading 0x{0:X}", iReturnCode);
                }
                else
                {
                    return resultDeviceValue;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            return -1;
        }

        int readPlcModel_M11()
        {
            int resultDeviceValue = 0;

            try
            {
                int iReturnCode = this.axActProgType1.GetDevice("M11", out resultDeviceValue);
                if (iReturnCode != 0)
                {
                    //  Error
                    Console.WriteLine("Error while reading 0x{0:X}", iReturnCode);
                }
                else
                {
                    return resultDeviceValue;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            return -1;
        }

        int readPlcModel_M12()
        {
            int resultDeviceValue = 0;

            try
            {
                int iReturnCode = this.axActProgType1.GetDevice("M12", out resultDeviceValue);
                if (iReturnCode != 0)
                {
                    //  Error
                    Console.WriteLine("Error while reading 0x{0:X}", iReturnCode);
                }
                else
                {
                    return resultDeviceValue;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            return -1;
        }
        
        int readPlcModel_M13()
        {
            int resultDeviceValue = 0;

            try
            {
                int iReturnCode = this.axActProgType1.GetDevice("M13", out resultDeviceValue);
                if (iReturnCode != 0)
                {
                    //  Error
                    Console.WriteLine("Error while reading 0x{0:X}", iReturnCode);
                }
                else
                {
                    return resultDeviceValue;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            return -1;
        }

        int readPlcModel_M14()
        {
            int resultDeviceValue = 0;

            try
            {
                int iReturnCode = this.axActProgType1.GetDevice("M14", out resultDeviceValue);
                if (iReturnCode != 0)
                {
                    //  Error
                    Console.WriteLine("Error while reading 0x{0:X}", iReturnCode);
                }
                else
                {
                    return resultDeviceValue;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            return -1;
        }

        int readPlcModel_M15()
        {
            int resultDeviceValue = 0;

            try
            {
                int iReturnCode = this.axActProgType1.GetDevice("M15", out resultDeviceValue);
                if (iReturnCode != 0)
                {
                    //  Error
                    Console.WriteLine("Error while reading 0x{0:X}", iReturnCode);
                }
                else
                {
                    return resultDeviceValue;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            return -1;
        }

        int readPlcModel_M16()
        {
            int resultDeviceValue = 0;

            try
            {
                int iReturnCode = this.axActProgType1.GetDevice("M16", out resultDeviceValue);
                if (iReturnCode != 0)
                {
                    //  Error
                    Console.WriteLine("Error while reading 0x{0:X}", iReturnCode);
                }
                else
                {
                    return resultDeviceValue;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            return -1;
        }

        int readPlcModel_M17()
        {
            int resultDeviceValue = 0;

            try
            {
                int iReturnCode = this.axActProgType1.GetDevice("M17", out resultDeviceValue);
                if (iReturnCode != 0)
                {
                    //  Error
                    Console.WriteLine("Error while reading 0x{0:X}", iReturnCode);
                }
                else
                {
                    return resultDeviceValue;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            return -1;
        }

        int readPlcModel_M18()
        {
            int resultDeviceValue = 0;

            try
            {
                int iReturnCode = this.axActProgType1.GetDevice("M18", out resultDeviceValue);
                if (iReturnCode != 0)
                {
                    //  Error
                    Console.WriteLine("Error while reading 0x{0:X}", iReturnCode);
                }
                else
                {
                    return resultDeviceValue;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            return -1;
        }

        int readPlcModel_M19()
        {
            int resultDeviceValue = 0;

            try
            {
                int iReturnCode = this.axActProgType1.GetDevice("M19", out resultDeviceValue);
                if (iReturnCode != 0)
                {
                    //  Error
                    Console.WriteLine("Error while reading 0x{0:X}", iReturnCode);
                }
                else
                {
                    return resultDeviceValue;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            return -1;
        }

        #endregion

        #region Handheld Events
        /************************************************************************/
        /*                           Hand-held Events                           */
        /************************************************************************/
        private void OnSerialSysDiscovered(SerSystemDiscoverer.SystemInfo sysInfo)
        {
            syncContext.Post(
                new SendOrPostCallback(
                    delegate
                    {
                        cbbDmrList.Items.Clear();
                        if (!cbbDmrList.Items.Contains(sysInfo)) //  Prevent add duplicate item
                        {
                            cbbDmrList.Items.Add(sysInfo);
                            cbbDmrList.SelectedIndex = cbbDmrList.FindStringExact(sysInfo.PortName);
                        }


                        foreach (String baudStr in baudArray)
                        {
                            if (!cbbDmrBaud.Items.Contains(baudStr.ToString()))
                            {
                                cbbDmrBaud.Items.Add(baudStr);
                            }
                        }

                        cbbDmrBaud.SelectedIndex = formInfo.baudIndex;                        
                    }), null);
        }
     
        private void OnEthSysDiscovered(EthSystemDiscoverer.SystemInfo sysInfo)
        {
            syncContext.Post(
                new SendOrPostCallback(
                    delegate
                    {
                        cbbDmrList.Items.Clear();
                        if (!cbbDmrList.Items.Contains(sysInfo)) //  Prevent add duplicate item
                        {
                            cbbDmrList.Items.Add(sysInfo);
                            cbbDmrList.SelectedIndex = cbbDmrList.FindStringExact(sysInfo.Name);
                        }
                    }), null);
        }

        private void OnSystemConnected(object sender, EventArgs args)
        {
            syncContext.Post(
                delegate
                {
                    //  Thread Connected
                    this.SerialPortState = (int)PortState.Connected_;
                    this.displayDMRportState();

                }, null);
        }

        private void OnSystemDisconnected(object sender, EventArgs args)
        {
            syncContext.Post(
                delegate
                {
                    //  Thread Disconnected
                    this.SerialPortState = (int)PortState.Disconnected_;
                    this.displayDMRportState();
                }, null);
        }

        private void OnComplexResultArrived(object sender, ResultInfo e)
        {
            syncContext.Post(
                delegate
                {
                    //  Thread Results Arrived
                    String tmpBarcodeStr = !String.IsNullOrEmpty(e.ReadString) ? e.ReadString : GetReadStringFromResultXml(e.XmlResult);
                    
                    

                    //  Prevent empty or null string.
                    if (tmpBarcodeStr.Length != Fixed_barcodeLen)
                    {
                        //  Null or wrong expected data.
                        
                    }
                    else
                    {
                        //  Replace spaces to "-".
                        tmpBarcodeStr = tmpBarcodeStr.Replace("      ", "-");

                        Console.WriteLine();
                        
                        Console.WriteLine("[{0}] Hand-held is read = true here!", DateTime.Now.ToString("H:mm:ss.fff"));

                        //if (String.Compare(tmpBarcodeStr, Results_HandheldNow, true) == 0)
                        if (String.Compare(tmpBarcodeStr, Results_HandheldNow, true) == 0)
                        {
                            //  Duplicate data received
                            Console.WriteLine("[{0}]Duplicate Handheld data capture: {1}", DateTime.Now.ToString("H:mm:ss.fff"), tmpBarcodeStr);
                        }
                        else
                        {
                            
                            
                            
                            //  New Bar-code Data Arrival
                            //  Queue it up!

                            

                            Console.WriteLine("[{0}] Goto Queue[{1}] | Lock: {2}", DateTime.Now.ToString("H:mm:ss.fff"), queuePosition, queueLock);


                            if (!queueLock)
                            {
                                queueLock = true;
                                if (queuePosition == 0)
                                {
                                    queuePosition = 1;
                                }
                                else
                                {
                                    queuePosition = 0;
                                }
                            }

                            Results_HandheldQueue[queuePosition] = tmpBarcodeStr;


                            Console.WriteLine("[{0}] Queue[0]: {1}", DateTime.Now.ToString("H:mm:ss.fff"), Results_HandheldQueue[0]);
                            Console.WriteLine("[{0}] Queue[1]: {1}", DateTime.Now.ToString("H:mm:ss.fff"), Results_HandheldQueue[1]);
                            Console.WriteLine("[{0}] DisplayString: {1}", DateTime.Now.ToString("H:mm:ss.fff"), displayHandheldString);
                            Console.WriteLine("[{0}] VisionRead {1} | QueueUp {2}", DateTime.Now.ToString("H:mm:ss.fff"), isVisionRead, isQueueUp);

                            if (VMSystemState)
                            {
                                if (isVisionRead)
                                {
                                    isVisionRead = false;

                                    if (isQueueUp)
                                    {
                                        Results_HandheldNow = Results_HandheldQueue[Math.Abs(1 - queuePosition)];
                                        Results_HandheldQueue[Math.Abs(1 - queuePosition)] = "";
                                    }
                                    else
                                    {
                                        Results_HandheldNow = Results_HandheldQueue[queuePosition];
                                    }
                                    

                                    isQueueUp = false;
                                }
                                else
                                {
                                    if (String.IsNullOrEmpty(displayHandheldString))
                                    {
                                        Results_HandheldNow = Results_HandheldQueue[queuePosition];
                                        queueLock = false;
                                    }
                                    else
                                    {
                                        if (!isQueueUp)
                                        {
                                            Results_HandheldNow = Results_HandheldQueue[Math.Abs(1- queuePosition)];
                                            isQueueUp = true;
                                            queueLock = false;
                                        }
                                    }

                                    
                                }
                            }

                            if (TMR_DisplayDelay.Enabled == false)
                            {
                                //  Reset GUI Components to default color.
                                displayResultReset();   //  Clear all format background
                                clearResultOK_NG(); //  Clear OK/NG background & data
                                clearResultLabel(false);    //  Clear all label except bar-code data.

                                

                                    Console.WriteLine("[{0}] Handheld Display Queue[{1}]: {2}", DateTime.Now.ToString("H:mm:ss.fff"), queuePosition, Results_HandheldNow);

                                    //  Append to DataTable
                                    addRowData();
                                    addBarcodeData(Results_HandheldNow);

                                    //  Displays
                                    setBarcodeText_v2(Results_HandheldNow);   // display on screen
                                    setTransferStr(Fixed_barcodeLen, Results_HandheldNow); //  Send to Camera via Modbus TCP
                                
                                
                            }
                            else
                            {

                            }
                        }

                    }

                    isHandheldRead = true;
                    //  Debug
                    //Console.WriteLine("temp: {0} | results: {1}", tmpBarcodeStr, Results_HandheldNow);

                }, null);
        }

        private void OnReadStringArrived(object sender, ReadStringArrivedEventArgs args)
        {
            Console.WriteLine(args.ReadString);
        }

        private void OnStatusEventArrived(object sender, StatusEventArrivedEventArgs args)
        {
            Console.WriteLine("Status Event Arrived {0}", args.ToString());
        }

        private void OnBinaryDataTransferProgress(object sender, BinaryDataTransferProgressEventArgs args)
        {
            syncContext.Post(
                delegate
                {
                    
                    double processValue = (int)(100 * (args.BytesTransferred / (double)args.TotalDataSize));
                    Console.WriteLine(processValue.ToString());
                },
                null);
        }

        

        #endregion

        #region Handheld Addition

        private void HandheldDiscover()
        {
            //  Discover Hand-held
            serialSysDiscovery = new SerSystemDiscoverer();
            ethSystemDiscoverer = new EthSystemDiscoverer();

            serialSysDiscovery.SystemDiscovered += new SerSystemDiscoverer.SystemDiscoveredHandler(OnSerialSysDiscovered);
            ethSystemDiscoverer.SystemDiscovered += new EthSystemDiscoverer.SystemDiscoveredHandler(OnEthSysDiscovered);

            serialSysDiscovery.Discover();
            ethSystemDiscoverer.Discover();
        }

        //  Read Hand-held XML String
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

        //  Cleanup Connection
        private void cleanupConnection()
        {
            if (dataManSys != null)
            {
                dataManSys.SystemConnected -= this.OnSystemConnected;
                dataManSys.SystemDisconnected -= this.OnSystemDisconnected;
                dataManSys.ReadStringArrived -= this.OnReadStringArrived;
                dataManSys.StatusEventArrived -= this.OnStatusEventArrived;
                dataManSys.BinaryDataTransferProgress -= this.OnBinaryDataTransferProgress;
            }
            sysConnector = null;
            dataManSys = null;
        }

        //  A-sync send command to Hand-Held (COGNEX Only)
        private void asyncBeginCmdSend(IAsyncResult results)
        {
//             Console.WriteLine();
//             Console.WriteLine("[{0}] Device on air: {1}", DateTime.Now.ToString("H:mm:ss.fff"), results.IsCompleted);

        }

        private void clickBtnDMRConnect()
        {
            this.btnDmrConnect.PerformClick();
        }

        private void clickBtnDMRDisconnect()
        {
            this.btnDmrDisconnect.PerformClick();
        }

        //  Send command every setup interval to check connection state
        private void Observe_Tick(object sender, EventArgs e)
        {
            //Console.WriteLine("[{0}] ### SerialPort State: {1}", DateTime.Now.ToString("H:mm:ss.fff"), SerialPortState_t[SerialPortState]);

            if (this.SerialPortState == (int)PortState.Disconnected_)
            {
                Console.WriteLine("[{0}] Connecting to device", DateTime.Now.ToString("H:mm:ss.fff"));
                syncContext.Post(new SendOrPostCallback(
                    delegate
                    {
                            this.clickBtnDMRConnect();
                    }), null);
            }

            if (this.SerialPortState == (int)PortState.ReConnecting_)
            {
                Console.WriteLine("[{0}] Reconnecting device", DateTime.Now.ToString("H:mm:ss.fff"));

                this.SerialPortState = (int)PortState.Connecting_;
                syncContext.Post(new SendOrPostCallback(
                    delegate{
                        this.clickBtnDMRDisconnect();
                    }), null);
                
            }

            if ((this.SerialPortState != (int)PortState.ReConnecting_) && (this.SerialPortState != (int)PortState.Connecting_))
            {
                try
                {
                    DMR_ResponseResults = dataManSys.EndSendCommand(dataManSys.BeginSendCommand("GET DEVICE.TYPE", asyncBeginCmdSend, dataManSys));

                    //Console.WriteLine("[{0}] Return value: {1}", DateTime.Now.ToString("H:mm:ss.fff"), DMR_ResponseResults.PayLoad);

                    if (this.SerialPortState == (int)PortState.LostConnection_)
                    {
                        this.SerialPortState = (int)PortState.ReConnecting_;
                    }

                }
                catch (Exception ex)
                {
                    inCorruptSerialConnection = true;
                    this.SerialPortState = (int)PortState.LostConnection_;
                    Console.WriteLine("[{0}] {1}", DateTime.Now.ToString("H:mm:ss.fff"), ex.Message);
                }
            }
            
            //Console.WriteLine("[{0}] Connection State: {1}", DateTime.Now.ToString("H:mm:ss.fff"), dataManSys.Connector.State);
            
        }
        #endregion

        #region Handheld SerialPort Data Transfer
        private void serialPort_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            try
            {
                RxString = serialPort.ReadLine();

            }
            catch (System.Exception ex)
            {
                if (serialPort.IsOpen)
                {
                    serialPort.Close();
                }
            }

            if (RxString != String.Empty)
            {
                Console.WriteLine(RxString);
                SetText(RxString);
            }


        }

        private void SetText(String inputText)
        {
            syncContext.Post(
                delegate
                {
                    if (inputText.Contains("|"))
                    {
                        

                        int startStrIndex = inputText.IndexOf("]");
                        String barcodeDataStr = inputText.Substring(startStrIndex + 1, inputText.Length - startStrIndex - 1);
                        Console.WriteLine("DMCC - advance COGNEX DataMan data transfer");
                        Console.WriteLine("DMCC Barcode Filtered: {0}{1}", barcodeDataStr, Environment.NewLine);
                    }
                    else
                    {
                        Console.WriteLine("Tradition RS-232 Transfer");


                        String tmpBarcodeStr = inputText;

                        if (tmpBarcodeStr.Length != Fixed_barcodeLen)
                        {
                            //  Null or wrong expected data.

                        }
                        else
                        {
                            isHandheldRead = true;
                            Console.WriteLine("[Tracer] Hand-held is read = true here!");

                            if (String.Compare(tmpBarcodeStr, Results_HandheldNow, true) == 0)
                            {
                                //  Duplicate data received
                                Console.WriteLine("Duplicate Handheld data capture: {0}", tmpBarcodeStr);
                            }
                            else
                            {
                                //  New data arrived
                                Results_HandheldNow = tmpBarcodeStr;


                                //  Reset GUI Components to default color.
                                displayResultReset();   //  Clear all format background
                                clearResultOK_NG(); //  Clear OK/NG background & data
                                clearResultLabel(false);    //  Clear all label except bar-code data.

                                if (VMSystemState)
                                {
                                    //  Append to DataTable
                                    addRowData();
                                    addBarcodeData(Results_HandheldNow);



                                    //  Displays
                                    setBarcodeText_v2(Results_HandheldNow);   // display on screen
                                    setTransferStr(Fixed_barcodeLen, Results_HandheldNow); //  Send to Camera via Modbus TCP
                                }

                                Console.WriteLine("New data arrived {0}", Results_HandheldNow);
                            }
                        }

                    }
                }, null);
        }
        #endregion

        #region ModbusTCP
        //  Transfer data to Server Buffer ---> ready for client read
        bool setTransferStr(Int16 StrLength, String strToSend)
        {
            Int16 registerPos_starStr = 3;
            Int16 registerPos_StrLength = 1;

            if (StrLength < strToSend.Length)
            {
                return false;
            }

            char[] charArr = strToSend.ToCharArray();

            mbServer.holdingRegisters[registerPos_StrLength] = StrLength;
            for (int i = 0; i < strToSend.Length; i++)
            {
                mbServer.holdingRegisters[registerPos_starStr + i] = Convert.ToByte(charArr[i]);
            }

            return true;
        }
        #endregion

        #region Model Control
        /************************************************************************/
        /*                        Add/Remove Model List                         */
        /************************************************************************/
        
        private void btn_AddModel_Click(object sender, EventArgs e)
        {
            if (txtModelInput.Text != String.Empty && txtJobInput.Text != String.Empty)
            {
                if (lbc_ModelList.Items.Contains(txtModelInput.Text))
                {

                }
                else
                {
                    lbc_ModelList.Items.Add(this.txtModelInput.Text);
                    lbc_JobList.Items.Add(this.txtJobInput.Text);
                    jobnames.Add(this.txtJobInput.Text);
                    modelnames.Add(this.txtModelInput.Text);
                    loaddone = 0;
                    //addToComboBox(lbc_ModelList, cbb_ModelList);
                    cbb_ModelList.DataSource = null;
                    cbb_ModelList.DataSource = modelnames;
                    loaddone = 1;
                    saveModel();
                    this.txtModelInput.Text = "";
                    this.txtJobInput.Text = "";
                }
            }
        }

        private void btn_RemoveModel_Click(object sender, EventArgs e)
        {
            index_select_before = -1;
            if (lbc_ModelList.SelectedIndex > -1)
            {
                DialogResult dialogResult = MessageBox.Show("Are you sure to delete this model?", "Delete Model", MessageBoxButtons.YesNoCancel);

                if (dialogResult == DialogResult.Yes)
                {
                    if (jobnames.Contains(txtJobInput.Text) && modelnames.Contains(txtModelInput.Text))
                    {
                        int index = lbc_ModelList.SelectedIndex;
                        lbc_JobList.Items.RemoveAt(index);
                        lbc_ModelList.Items.RemoveAt(index);
                        modelnames.RemoveAt(index);
                        jobnames.RemoveAt(index);
                        //this.Cursor = Cursors.Default;
                        txtModelInput.Text = "";
                        txtJobInput.Text = "";
                        //lbc_ModelList.DataSource = modelnames;
                        //lbc_JobList.DataSource = jobnames;
                        //addToComboBox(lbc_ModelList, cbb_ModelList);
                        cbb_ModelList.DataSource = null;
                        cbb_ModelList.DataSource = modelnames;

                    }
                    else
                    {
                        
                    }
                }
                else if (dialogResult == DialogResult.No)
                {
                    lbc_ModelList.Select();
                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    
                }
                
                
            }
        }

        private void saveModel()
        {
            formInfo.modelList = "";
            formInfo.jobList = "";
            //formInfo.CurrentModel = "";
            //for (int i = 0; i < lbc_ModelList.Items.Count; i++)
            //{
            //    formInfo.modelList += lbc_ModelList.Items[i].ToString() + ",";
            //}
            //Console.WriteLine("Saving Listbox {0}", formInfo.modelList);

            foreach (var model in modelnames)
            {
                formInfo.modelList += model + ",";
            }
            foreach (var job in jobnames)
            {
                formInfo.jobList += job + ",";
            }
            //formInfo.CurrentModel = index_select_before.ToString();
            Console.WriteLine("Saving Listbox {0}", formInfo.modelList);
        }

        private void loadModel()
        {
            //String[] modelString = formInfo.modelList.Split(',');
            //foreach (string item in modelString)
            //{
            //    if (item.Trim() == "")
            //    {
            //        continue;
            //    }
            //    lbc_ModelList.Items.Add(item);
            //    cbb_ModelList.Items.Add(item);
            //    cbb_ModelList.SelectedIndex = cbb_ModelList.Items.IndexOf(item);
            //    Console.WriteLine(item);
            //}

            String[] modelString = formInfo.modelList.Split(',');
            String[] jobString = formInfo.jobList.Split(',');
            foreach (string item in modelString)
            {
                if (item.Trim() == "")
                {
                    continue;
                }
                lbc_ModelList.Items.Add(item);
                cbb_ModelList.Items.Add(item);
                modelnames.Add(item);
                cbb_ModelList.SelectedIndex = cbb_ModelList.Items.IndexOf(item);
                Console.WriteLine(item);
            }
            foreach (string item in jobString)
            {
                if (item.Trim() == "")
                {
                    continue;
                }
                lbc_JobList.Items.Add(item);
                
                jobnames.Add(item);
                //lbljobfile.Text = jobnames[0];
                Console.WriteLine(item);
            }
            index_select_before = Convert.ToInt32(formInfo.CurrentModel);
            cbb_ModelList.SelectedIndex = index_select_before;
            //lbljobfile.Text = jobnames[modelnames.IndexOf(formInfo.CurrentModel)];
        }


        /************************************************************************/
        /*                        Control Model in List                         */
        /************************************************************************/
        private void addToComboBox(ListBox sourceListBox, ComboBox outComboBox)
        {
            for (int i = 0; i < sourceListBox.Items.Count; i++)
            {
                if (!outComboBox.Items.Contains(sourceListBox.Items[i]))
                {
                    //  Prevent add duplicate item
                    outComboBox.Items.Add(sourceListBox.Items[i]);
                }
            }
        }

        #endregion

        #region Dock Events

        private void dockManager_Expanding(object sender, DevExpress.XtraBars.Docking.DockPanelCancelEventArgs e)
        {
            //HandheldDiscover();
        }
        #endregion

        #region CharPropertiesDisplay
        /*
        private void inputCharSpecs(int charCol, double inputHeigh, double inputWidth, double inputAngle, double inputUpLimitHeight, double inputLowLimitHeight, double inputUpLimitWidth, double inputLowLimitWidth, double inputUpLimitAngle, double inputLowLimitAngle)
        {
            switch (charCol)
            {
                case 1:
                {
                    lbl_H01.Text = inputHeigh.ToString("F3");
                    lbl_W01.Text = inputWidth.ToString("F3");
                    lbl_A01.Text = inputAngle.ToString("F3");

                    if(!inputHeigh.InRange(inputLowLimitHeight, inputUpLimitHeight))
                    {
                        lbl_H01.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputWidth.InRange(inputLowLimitWidth, inputUpLimitWidth))
                    {
                        lbl_W01.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputAngle.InRange(inputLowLimitAngle, inputUpLimitAngle))
                    {
                        lbl_A01.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 2:
                {
                    lbl_H02.Text = inputHeigh.ToString("F3");
                    lbl_W02.Text = inputWidth.ToString("F3");
                    lbl_A02.Text = inputAngle.ToString("F3");

                    if (!inputHeigh.InRange(inputLowLimitHeight, inputUpLimitHeight))
                    {
                        lbl_H02.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputWidth.InRange(inputLowLimitWidth, inputUpLimitWidth))
                    {
                        lbl_W02.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputAngle.InRange(inputLowLimitAngle, inputUpLimitAngle))
                    {
                        lbl_A02.BackColor = System.Drawing.Color.Red;
                    }
                    break;
                }

                case 3:
                {
                    lbl_H03.Text = inputHeigh.ToString("F3");
                    lbl_W03.Text = inputWidth.ToString("F3");
                    lbl_A03.Text = inputAngle.ToString("F3");

                    if (!inputHeigh.InRange(inputLowLimitHeight, inputUpLimitHeight))
                    {
                        lbl_H03.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputWidth.InRange(inputLowLimitWidth, inputUpLimitWidth))
                    {
                        lbl_W03.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputAngle.InRange(inputLowLimitAngle, inputUpLimitAngle))
                    {
                        lbl_A03.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 4:
                {
                    lbl_H04.Text = inputHeigh.ToString("F3");
                    lbl_W04.Text = inputWidth.ToString("F3");
                    lbl_A04.Text = inputAngle.ToString("F3");

                    if (!inputHeigh.InRange(inputLowLimitHeight, inputUpLimitHeight))
                    {
                        lbl_H04.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputWidth.InRange(inputLowLimitWidth, inputUpLimitWidth))
                    {
                        lbl_W04.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputAngle.InRange(inputLowLimitAngle, inputUpLimitAngle))
                    {
                        lbl_A04.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 5:
                {
                    lbl_H05.Text = inputHeigh.ToString("F3");
                    lbl_W05.Text = inputWidth.ToString("F3");
                    lbl_A05.Text = inputAngle.ToString("F3");

                    if (!inputHeigh.InRange(inputLowLimitHeight, inputUpLimitHeight))
                    {
                        lbl_H05.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputWidth.InRange(inputLowLimitWidth, inputUpLimitWidth))
                    {
                        lbl_W05.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputAngle.InRange(inputLowLimitAngle, inputUpLimitAngle))
                    {
                        lbl_A05.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 6:
                {
                    lbl_H06.Text = inputHeigh.ToString("F3");
                    lbl_W06.Text = inputWidth.ToString("F3");
                    lbl_A06.Text = inputAngle.ToString("F3");

                    if (!inputHeigh.InRange(inputLowLimitHeight, inputUpLimitHeight))
                    {
                        lbl_H06.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputWidth.InRange(inputLowLimitWidth, inputUpLimitWidth))
                    {
                        lbl_W06.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputAngle.InRange(inputLowLimitAngle, inputUpLimitAngle))
                    {
                        lbl_A06.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 7:
                {
                    lbl_H07.Text = inputHeigh.ToString("F3");
                    lbl_W07.Text = inputWidth.ToString("F3");
                    lbl_A07.Text = inputAngle.ToString("F3");

                    if (!inputHeigh.InRange(inputLowLimitHeight, inputUpLimitHeight))
                    {
                        lbl_H07.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputWidth.InRange(inputLowLimitWidth, inputUpLimitWidth))
                    {
                        lbl_W07.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputAngle.InRange(inputLowLimitAngle, inputUpLimitAngle))
                    {
                        lbl_A07.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 8:
                {
                    lbl_H08.Text = inputHeigh.ToString("F3");
                    lbl_W08.Text = inputWidth.ToString("F3");
                    lbl_A08.Text = inputAngle.ToString("F3");

                    if (!inputHeigh.InRange(inputLowLimitHeight, inputUpLimitHeight))
                    {
                        lbl_H08.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputWidth.InRange(inputLowLimitWidth, inputUpLimitWidth))
                    {
                        lbl_W08.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputAngle.InRange(inputLowLimitAngle, inputUpLimitAngle))
                    {
                        lbl_A08.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 9:
                {
                    lbl_H09.Text = inputHeigh.ToString("F3");
                    lbl_W09.Text = inputWidth.ToString("F3");
                    lbl_A09.Text = inputAngle.ToString("F3");

                    if (!inputHeigh.InRange(inputLowLimitHeight, inputUpLimitHeight))
                    {
                        lbl_H09.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputWidth.InRange(inputLowLimitWidth, inputUpLimitWidth))
                    {
                        lbl_W09.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputAngle.InRange(inputLowLimitAngle, inputUpLimitAngle))
                    {
                        lbl_A09.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 10:
                {
                    lbl_H10.Text = inputHeigh.ToString("F3");
                    lbl_W10.Text = inputWidth.ToString("F3");
                    lbl_A10.Text = inputAngle.ToString("F3");

                    if (!inputHeigh.InRange(inputLowLimitHeight, inputUpLimitHeight))
                    {
                        lbl_H10.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputWidth.InRange(inputLowLimitWidth, inputUpLimitWidth))
                    {
                        lbl_W10.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputAngle.InRange(inputLowLimitAngle, inputUpLimitAngle))
                    {
                        lbl_A10.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 11:
                {
                    lbl_H11.Text = inputHeigh.ToString("F3");
                    lbl_W11.Text = inputWidth.ToString("F3");
                    lbl_A11.Text = inputAngle.ToString("F3");

                    if (!inputHeigh.InRange(inputLowLimitHeight, inputUpLimitHeight))
                    {
                        lbl_H11.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputWidth.InRange(inputLowLimitWidth, inputUpLimitWidth))
                    {
                        lbl_W11.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputAngle.InRange(inputLowLimitAngle, inputUpLimitAngle))
                    {
                        lbl_A11.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 12:
                {
                    lbl_H12.Text = inputHeigh.ToString("F3");
                    lbl_W12.Text = inputWidth.ToString("F3");
                    lbl_A12.Text = inputAngle.ToString("F3");

                    if (!inputHeigh.InRange(inputLowLimitHeight, inputUpLimitHeight))
                    {
                        lbl_H12.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputWidth.InRange(inputLowLimitWidth, inputUpLimitWidth))
                    {
                        lbl_W12.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputAngle.InRange(inputLowLimitAngle, inputUpLimitAngle))
                    {
                        lbl_A12.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 13:
                {
                    lbl_H13.Text = inputHeigh.ToString("F3");
                    lbl_W13.Text = inputWidth.ToString("F3");
                    lbl_A13.Text = inputAngle.ToString("F3");

                    if (!inputHeigh.InRange(inputLowLimitHeight, inputUpLimitHeight))
                    {
                        lbl_H13.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputWidth.InRange(inputLowLimitWidth, inputUpLimitWidth))
                    {
                        lbl_W13.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputAngle.InRange(inputLowLimitAngle, inputUpLimitAngle))
                    {
                        lbl_A13.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 14:
                {
                    lbl_H14.Text = inputHeigh.ToString("F3");
                    lbl_W14.Text = inputWidth.ToString("F3");
                    lbl_A14.Text = inputAngle.ToString("F3");

                    if (!inputHeigh.InRange(inputLowLimitHeight, inputUpLimitHeight))
                    {
                        lbl_H14.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputWidth.InRange(inputLowLimitWidth, inputUpLimitWidth))
                    {
                        lbl_W14.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputAngle.InRange(inputLowLimitAngle, inputUpLimitAngle))
                    {
                        lbl_A14.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 15:
                {
                    lbl_H15.Text = inputHeigh.ToString("F3");
                    lbl_W15.Text = inputWidth.ToString("F3");
                    lbl_A15.Text = inputAngle.ToString("F3");

                    if (!inputHeigh.InRange(inputLowLimitHeight, inputUpLimitHeight))
                    {
                        lbl_H15.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputWidth.InRange(inputLowLimitWidth, inputUpLimitWidth))
                    {
                        lbl_W15.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputAngle.InRange(inputLowLimitAngle, inputUpLimitAngle))
                    {
                        lbl_A15.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 16:
                {
                    lbl_H16.Text = inputHeigh.ToString("F3");
                    lbl_W16.Text = inputWidth.ToString("F3");
                    lbl_A16.Text = inputAngle.ToString("F3");

                    if (!inputHeigh.InRange(inputLowLimitHeight, inputUpLimitHeight))
                    {
                        lbl_H16.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputWidth.InRange(inputLowLimitWidth, inputUpLimitWidth))
                    {
                        lbl_W16.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputAngle.InRange(inputLowLimitAngle, inputUpLimitAngle))
                    {
                        lbl_A16.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 17:
                {
                    lbl_H17.Text = inputHeigh.ToString("F3");
                    lbl_W17.Text = inputWidth.ToString("F3");
                    lbl_A17.Text = inputAngle.ToString("F3");


                    if (!inputHeigh.InRange(inputLowLimitHeight, inputUpLimitHeight))
                    {
                        lbl_H17.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputWidth.InRange(inputLowLimitWidth, inputUpLimitWidth))
                    {
                        lbl_W17.BackColor = System.Drawing.Color.Red;
                    }

                    if (!inputAngle.InRange(inputLowLimitAngle, inputUpLimitAngle))
                    {
                        lbl_A17.BackColor = System.Drawing.Color.Red;
                    }


                    break;
                }
            }
        }
        */

        private void inputCharSpecs(int charCol, double inputHeight, double inputWidth, double inputAngle, bool heightOK, bool widthOK, bool angleOK)
        {
            switch (charCol)
            {
                case 1:
                {
                    lbl_H01.Text = inputHeight.ToString("F3");
                    lbl_W01.Text = inputWidth.ToString("F3");
                    lbl_A01.Text = inputAngle.ToString("F3");

                    if (checkList[0] && !heightOK)
                    {
                        lbl_H01.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[1] && !widthOK)
                    {
                        lbl_W01.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[2] && !angleOK)
                    {
                        lbl_A01.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 2:
                {
                    lbl_H02.Text = inputHeight.ToString("F3");
                    lbl_W02.Text = inputWidth.ToString("F3");
                    lbl_A02.Text = inputAngle.ToString("F3");

                    if (checkList[0] && !heightOK)
                    {
                        lbl_H02.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[1] && !widthOK)
                    {
                        lbl_W02.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[2] && !angleOK)
                    {
                        lbl_A02.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 3:
                {
                    lbl_H03.Text = inputHeight.ToString("F3");
                    lbl_W03.Text = inputWidth.ToString("F3");
                    lbl_A03.Text = inputAngle.ToString("F3");

                    if (checkList[0] && !heightOK)
                    {
                        lbl_H03.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[1] && !widthOK)
                    {
                        lbl_W03.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[2] && !angleOK)
                    {
                        lbl_A03.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 4:
                {
                    lbl_H04.Text = inputHeight.ToString("F3");
                    lbl_W04.Text = inputWidth.ToString("F3");
                    lbl_A04.Text = inputAngle.ToString("F3");

                    if (checkList[0] && !heightOK)
                    {
                        lbl_H04.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[1] && !widthOK)
                    {
                        lbl_W04.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[2] && !angleOK)
                    {
                        lbl_A04.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 5:
                {
                    lbl_H05.Text = inputHeight.ToString("F3");
                    lbl_W05.Text = inputWidth.ToString("F3");
                    lbl_A05.Text = inputAngle.ToString("F3");

                    if (checkList[0] && !heightOK)
                    {
                        lbl_H05.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[1] && !widthOK)
                    {
                        lbl_W05.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[2] && !angleOK)
                    {
                        lbl_A05.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 6:
                {
                    lbl_H06.Text = inputHeight.ToString("F3");
                    lbl_W06.Text = inputWidth.ToString("F3");
                    lbl_A06.Text = inputAngle.ToString("F3");

                    if (checkList[0] && !heightOK)
                    {
                        lbl_H06.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[1] && !widthOK)
                    {
                        lbl_W06.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[2] && !angleOK)
                    {
                        lbl_A06.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 7:
                {
                    lbl_H07.Text = inputHeight.ToString("F3");
                    lbl_W07.Text = inputWidth.ToString("F3");
                    lbl_A07.Text = inputAngle.ToString("F3");

                    if (checkList[0] && !heightOK)
                    {
                        lbl_H07.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[1] && !widthOK)
                    {
                        lbl_W07.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[2] && !angleOK)
                    {
                        lbl_A07.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 8:
                {
                    lbl_H08.Text = inputHeight.ToString("F3");
                    lbl_W08.Text = inputWidth.ToString("F3");
                    lbl_A08.Text = inputAngle.ToString("F3");

                    if (checkList[0] && !heightOK)
                    {
                        lbl_H08.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[1] && !widthOK)
                    {
                        lbl_W08.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[2] && !angleOK)
                    {
                        lbl_A08.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 9:
                {
                    lbl_H09.Text = inputHeight.ToString("F3");
                    lbl_W09.Text = inputWidth.ToString("F3");
                    lbl_A09.Text = inputAngle.ToString("F3");

                    if (checkList[0] && !heightOK)
                    {
                        lbl_H09.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[1] && !widthOK)
                    {
                        lbl_W09.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[2] && !angleOK)
                    {
                        lbl_A09.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 10:
                {
                    lbl_H10.Text = inputHeight.ToString("F3");
                    lbl_W10.Text = inputWidth.ToString("F3");
                    lbl_A10.Text = inputAngle.ToString("F3");

                    if (checkList[0] && !heightOK)
                    {
                        lbl_H10.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[1] && !widthOK)
                    {
                        lbl_W10.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[2] && !angleOK)
                    {
                        lbl_A10.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 11:
                {
                    lbl_H11.Text = inputHeight.ToString("F3");
                    lbl_W11.Text = inputWidth.ToString("F3");
                    lbl_A11.Text = inputAngle.ToString("F3");

                    if (checkList[0] && !heightOK)
                    {
                        lbl_H11.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[1] && !widthOK)
                    {
                        lbl_W11.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[2] && !angleOK)
                    {
                        lbl_A11.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 12:
                {
                    lbl_H12.Text = inputHeight.ToString("F3");
                    lbl_W12.Text = inputWidth.ToString("F3");
                    lbl_A12.Text = inputAngle.ToString("F3");

                    if (checkList[0] && !heightOK)
                    {
                        lbl_H12.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[1] && !widthOK)
                    {
                        lbl_W12.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[2] && !angleOK)
                    {
                        lbl_A12.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 13:
                {
                    lbl_H13.Text = inputHeight.ToString("F3");
                    lbl_W13.Text = inputWidth.ToString("F3");
                    lbl_A13.Text = inputAngle.ToString("F3");

                    if (checkList[0] && !heightOK)
                    {
                        lbl_H13.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[1] && !widthOK)
                    {
                        lbl_W13.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[2] && !angleOK)
                    {
                        lbl_A13.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 14:
                {
                    lbl_H14.Text = inputHeight.ToString("F3");
                    lbl_W14.Text = inputWidth.ToString("F3");
                    lbl_A14.Text = inputAngle.ToString("F3");

                    if (checkList[0] && !heightOK)
                    {
                        lbl_H14.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[1] && !widthOK)
                    {
                        lbl_W14.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[2] && !angleOK)
                    {
                        lbl_A14.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 15:
                {
                    lbl_H15.Text = inputHeight.ToString("F3");
                    lbl_W15.Text = inputWidth.ToString("F3");
                    lbl_A15.Text = inputAngle.ToString("F3");

                    if (checkList[0] && !heightOK)
                    {
                        lbl_H15.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[1] && !widthOK)
                    {
                        lbl_W15.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[2] && !angleOK)
                    {
                        lbl_A15.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 16:
                {
                    lbl_H16.Text = inputHeight.ToString("F3");
                    lbl_W16.Text = inputWidth.ToString("F3");
                    lbl_A16.Text = inputAngle.ToString("F3");

                    if (checkList[0] && !heightOK)
                    {
                        lbl_H16.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[1] && !widthOK)
                    {
                        lbl_W16.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[2] && !angleOK)
                    {
                        lbl_A16.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 17:
                {
                    lbl_H17.Text = inputHeight.ToString("F3");
                    lbl_W17.Text = inputWidth.ToString("F3");
                    lbl_A17.Text = inputAngle.ToString("F3");

                    if (checkList[0] && !heightOK)
                    {
                        lbl_H17.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[1] && !widthOK)
                    {
                        lbl_W17.BackColor = System.Drawing.Color.Red;
                    }

                    if (checkList[2] && !angleOK)
                    {
                        lbl_A17.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }
            }
        }

        /*
        private void inputDistanceCenter2Center(int pairOfCol, double inputDistance, double inputUpLimit, double inputLowLimit)
        {
            switch (pairOfCol)
            {
                case 1:
                {
                    lbl_C01.Text = inputDistance.ToString("F3");

                    if (!inputDistance.InRange(inputLowLimit, inputUpLimit))
                    {
                        lbl_C01.BackColor = System.Drawing.Color.Red;
                    }
                    break;
                }

                case 2:
                {
                    lbl_C02.Text = inputDistance.ToString("F3");

                    if (!inputDistance.InRange(inputLowLimit, inputUpLimit))
                    {
                        lbl_C02.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 3:
                {
                    lbl_C03.Text = inputDistance.ToString("F3");

                    if (!inputDistance.InRange(inputLowLimit, inputUpLimit))
                    {
                        lbl_C03.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 4:
                {
                    lbl_C04.Text = inputDistance.ToString("F3");

                    if (!inputDistance.InRange(inputLowLimit, inputUpLimit))
                    {
                        lbl_C04.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 5:
                {
                    lbl_C05.Text = inputDistance.ToString("F3");

                    if (!inputDistance.InRange(inputLowLimit, inputUpLimit))
                    {
                        lbl_C05.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 6:
                {
                    lbl_C06.Text = inputDistance.ToString("F3");

                    if (!inputDistance.InRange(inputLowLimit, inputUpLimit))
                    {
                        lbl_C06.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 7:
                {
                    lbl_C07.Text = inputDistance.ToString("F3");

                    if (!inputDistance.InRange(inputLowLimit, inputUpLimit))
                    {
                        lbl_C07.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 8:
                {
                    lbl_C08.Text = inputDistance.ToString("F3");

                    if (!inputDistance.InRange(inputLowLimit, inputUpLimit))
                    {
                        lbl_C08.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }
                
                case 9:
                {
                    lbl_C09.Text = inputDistance.ToString("F3");

                    if (!inputDistance.InRange(inputLowLimit, inputUpLimit))
                    {
                        lbl_C09.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 10:
                {
                    lbl_C10.Text = inputDistance.ToString("F3");
                    
                    if (!inputDistance.InRange(inputLowLimit, inputUpLimit))
                    {
                        lbl_C10.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 11:
                {
                    lbl_C11.Text = inputDistance.ToString("F3");

                    if (!inputDistance.InRange(inputLowLimit, inputUpLimit))
                    {
                        lbl_C11.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 12:
                {
                    lbl_C12.Text = inputDistance.ToString("F3");

                    if (!inputDistance.InRange(inputLowLimit, inputUpLimit))
                    {
                        lbl_C12.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 13:
                {
                    lbl_C13.Text = inputDistance.ToString("F3");

                    if (!inputDistance.InRange(inputLowLimit, inputUpLimit))
                    {
                        lbl_C13.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 14:
                {
                    lbl_C14.Text = inputDistance.ToString("F3");

                    if (!inputDistance.InRange(inputLowLimit, inputUpLimit))
                    {
                        lbl_C14.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }
                case 15:
                {
                    lbl_C15.Text = inputDistance.ToString("F3");

                    if (!inputDistance.InRange(inputLowLimit, inputUpLimit))
                    {
                        lbl_C15.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 16:
                {
                    lbl_C16.Text = inputDistance.ToString("F3");

                    if (!inputDistance.InRange(inputLowLimit, inputUpLimit))
                    {
                        lbl_C16.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }
                
            }
        }
        */
        private void inputDistanceCenter2Center(int pairOfCol, double inputDistance, bool distanceOK)
        {
            switch (pairOfCol)
            {
                case 1:
                {
                    lbl_C01.Text = inputDistance.ToString("F3");

                    if (checkList[3] && !distanceOK)
                    {
                        lbl_C01.BackColor = System.Drawing.Color.Red;
                    }
                    break;
                }

                case 2:
                {
                    lbl_C02.Text = inputDistance.ToString("F3");

                    if (checkList[3] && !distanceOK)
                    {
                        lbl_C02.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 3:
                {
                    lbl_C03.Text = inputDistance.ToString("F3");

                    if (checkList[3] && !distanceOK)
                    {
                        lbl_C03.BackColor = System.Drawing.Color.Red;
                    }

                    break;
                }

                case 4:
                    {
                        lbl_C04.Text = inputDistance.ToString("F3");

                        if (checkList[3] && !distanceOK)
                        {
                            lbl_C04.BackColor = System.Drawing.Color.Red;
                        }

                        break;
                    }

                case 5:
                    {
                        lbl_C05.Text = inputDistance.ToString("F3");

                        if (checkList[3] && !distanceOK)
                        {
                            lbl_C05.BackColor = System.Drawing.Color.Red;
                        }

                        break;
                    }

                case 6:
                    {
                        lbl_C06.Text = inputDistance.ToString("F3");

                        if (checkList[3] && !distanceOK)
                        {
                            lbl_C06.BackColor = System.Drawing.Color.Red;
                        }

                        break;
                    }

                case 7:
                    {
                        lbl_C07.Text = inputDistance.ToString("F3");

                        if (checkList[3] && !distanceOK)
                        {
                            lbl_C07.BackColor = System.Drawing.Color.Red;
                        }

                        break;
                    }

                case 8:
                    {
                        lbl_C08.Text = inputDistance.ToString("F3");

                        if (checkList[3] && !distanceOK)
                        {
                            lbl_C08.BackColor = System.Drawing.Color.Red;
                        }

                        break;
                    }

                case 9:
                    {
                        lbl_C09.Text = inputDistance.ToString("F3");

                        if (checkList[3] && !distanceOK)
                        {
                            lbl_C09.BackColor = System.Drawing.Color.Red;
                        }

                        break;
                    }

                case 10:
                    {
                        lbl_C10.Text = inputDistance.ToString("F3");

                        if (checkList[3] && !distanceOK)
                        {
                            lbl_C10.BackColor = System.Drawing.Color.Red;
                        }

                        break;
                    }

                case 11:
                    {
                        lbl_C11.Text = inputDistance.ToString("F3");

                        if (checkList[3] && !distanceOK)
                        {
                            lbl_C11.BackColor = System.Drawing.Color.Red;
                        }

                        break;
                    }

                case 12:
                    {
                        lbl_C12.Text = inputDistance.ToString("F3");

                        if (checkList[3] && !distanceOK)
                        {
                            lbl_C12.BackColor = System.Drawing.Color.Red;
                        }

                        break;
                    }

                case 13:
                    {
                        lbl_C13.Text = inputDistance.ToString("F3");

                        if (checkList[3] && !distanceOK)
                        {
                            lbl_C13.BackColor = System.Drawing.Color.Red;
                        }

                        break;
                    }

                case 14:
                    {
                        lbl_C14.Text = inputDistance.ToString("F3");

                        if (checkList[3] && !distanceOK)
                        {
                            lbl_C14.BackColor = System.Drawing.Color.Red;
                        }

                        break;
                    }
                case 15:
                    {
                        lbl_C15.Text = inputDistance.ToString("F3");

                        if (checkList[3] && !distanceOK)
                        {
                            lbl_C15.BackColor = System.Drawing.Color.Red;
                        }

                        break;
                    }

                case 16:
                    {
                        lbl_C16.Text = inputDistance.ToString("F3");

                        if (checkList[3] && !distanceOK)
                        {
                            lbl_C16.BackColor = System.Drawing.Color.Red;
                        }

                        break;
                    }

            }
        }

        /*
        private void inputLineStrAngle(double inputAngle, double lowLimitAngle, double upLimitAngle)
        {
            lbl_LA01.Text = inputAngle.ToString("F3");

            if (!inputAngle.InRange(lowLimitAngle, upLimitAngle))
            {
                lbl_LA01.BackColor = System.Drawing.Color.Red;
            }
            
        }
        */
        private void inputLineStrAngle(double inputAngle, bool lineAngleOK)
        {
            lbl_LA01.Text = inputAngle.ToString("F3");

            if (checkList[4] && !lineAngleOK)
            {
                lbl_LA01.BackColor = System.Drawing.Color.Red;
            }

        }
        #endregion

        #region SpreadSheet
        private void initTable()
        {
            //  Init DataTable
            resultTable.Columns.Add(tableCols[(int)ColPosition.Col_No], typeof(int));
            resultTable.Columns.Add(tableCols[(int)ColPosition.Col_Date], typeof(DateTime));
            resultTable.Columns.Add(tableCols[(int)ColPosition.Col_Time], typeof(DateTime));
            resultTable.Columns.Add(tableCols[(int)ColPosition.Col_Shift], typeof(String));
            resultTable.Columns.Add(tableCols[(int)ColPosition.Col_Model], typeof(String));
            resultTable.Columns.Add(tableCols[(int)ColPosition.Col_BarcodeData], typeof(String));
            resultTable.Columns.Add(tableCols[(int)ColPosition.Col_VisionData], typeof(String));
            resultTable.Columns.Add(tableCols[(int)ColPosition.Col_Result], typeof(String));

            dgvResultTable.DataSource = resultTable;
            dgvResultTable.Columns[(int)ColPosition.Col_Date].DefaultCellStyle.Format = "dd-MMM-yyyy";
            dgvResultTable.Columns[(int)ColPosition.Col_Time].DefaultCellStyle.Format = "HH:mm:ss";
            
        }
        
        private void scrollToEOL()
        {
            dgvResultTable.FirstDisplayedScrollingRowIndex = dgvResultTable.RowCount - 1;
        }

        private void addBarcodeData(String codeStr)
        {
            resultTable.Rows[resultTable.Rows.Count - 1][(int)ColPosition.Col_BarcodeData] = codeStr;
        }

        private void addVisionData(String visionStr)
        {
            resultTable.Rows[resultTable.Rows.Count - 1][(int)ColPosition.Col_VisionData] = visionStr;
        }

        private void addResult_OK_NG(bool inputResult)
        {
            if (inputResult)
            {
                resultTable.Rows[resultTable.Rows.Count - 1][(int)ColPosition.Col_Result] = "OK";
            }
            else
            {
                resultTable.Rows[resultTable.Rows.Count - 1][(int)ColPosition.Col_Result] = "NG";
            }
            
        }

        private void addShift(String shiftStr)
        {
            resultTable.Rows[resultTable.Rows.Count - 1][(int)ColPosition.Col_Shift] = shiftStr;
        }

        private void addModel(String modelStr)
        {
            resultTable.Rows[resultTable.Rows.Count - 1][(int)ColPosition.Col_Model] = modelStr;
        }

        private void addNowDate(String nowDate)
        {
            resultTable.Rows[resultTable.Rows.Count - 1][(int)ColPosition.Col_Date] = nowDate;
        }

        private void addNowTime(String nowTime)
        {
            resultTable.Rows[resultTable.Rows.Count - 1][(int)ColPosition.Col_Time] = nowTime;
        }

        private void addRowData()
        {
            DataRow tableRows = resultTable.NewRow();
            tableRows[tableCols[0]] = resultTable.Rows.Count + 1;
            tableRows[tableCols[1]] = DateTime.Now.ToString("dd-MMM-yyyy");
            tableRows[tableCols[2]] = DateTime.Now.ToString("H:mm:ss");
            tableRows[tableCols[3]] = cbb_Shift.SelectedItem.ToString();
            tableRows[tableCols[4]] = cbb_ModelList.SelectedItem.ToString();
            resultTable.Rows.Add(tableRows);
            scrollToEOL();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Excel Workbook|*.xlsx";
            saveFileDialog1.Title = "Save Report File";
            saveFileDialog1.FileName = DateTime.Now.ToString("dd.MM.yyyy_hhmmss");

            DialogResult results = saveFileDialog1.ShowDialog();

            if (results == DialogResult.OK)
            {
                System.IO.FileStream fs = (System.IO.FileStream)saveFileDialog1.OpenFile();
                SLDocument slDocs = new SLDocument();

                slDocs.ImportDataTable(1, 1, resultTable, true);
                SLStyle slStyle = slDocs.CreateStyle();
                slStyle.FormatCode = "dd-MMM-yyyy";
                slDocs.SetColumnStyle((int)ColPosition.Col_Date + 1, slStyle);
                slStyle.FormatCode = "HH:mm:ss";
                slDocs.SetColumnStyle((int)ColPosition.Col_Time + 1, slStyle);
                slDocs.SaveAs(fs);

                fs.Close();
            }

        }

        private void saveSpreadsheet()
        {
            String fullFilePath = LogFilePath + "\\" + DateTime.Now.ToString(LogFileName);
            String fullFileName = fullFilePath + ".xlsx";
            
            Console.WriteLine(fullFilePath);
            
            SLDocument slDocs = new SLDocument();
            slDocs.ImportDataTable(1, 1, resultTable, true);
            SLStyle slStyle = slDocs.CreateStyle();
            slStyle.FormatCode = "dd-MMM-yyyy";
            slDocs.SetColumnStyle((int)ColPosition.Col_Date + 1, slStyle);
            slStyle.FormatCode = "HH:mm:ss";
            slDocs.SetColumnStyle((int)ColPosition.Col_Time + 1, slStyle);
            slDocs.SaveAs(fullFileName);
        }

        private void loadSpreadsheet()
        {
            String fullFilePath = LogFilePath + "\\" + DateTime.Now.ToString(LogFileName);

            String fullFileName = fullFilePath + ".xlsx";


            if (System.IO.File.Exists(fullFileName))
            {
                using (SLDocument slDocs = new SLDocument(fullFileName, "Sheet1"))
                {
                    SLWorksheetStatistics stats = slDocs.GetWorksheetStatistics();
                    int iStartColumnIndex = stats.StartColumnIndex;

                    for (int i = stats.StartRowIndex + 1; i <= stats.EndRowIndex; ++i)
                    {
                        DataRow tableRows = resultTable.NewRow();
                        tableRows[tableCols[0]] = slDocs.GetCellValueAsInt32(i, iStartColumnIndex);
                        tableRows[tableCols[1]] = slDocs.GetCellValueAsDateTime(i, iStartColumnIndex + 1);
                        tableRows[tableCols[2]] = slDocs.GetCellValueAsDateTime(i, iStartColumnIndex + 2);
                        tableRows[tableCols[3]] = slDocs.GetCellValueAsString(i, iStartColumnIndex + 3);
                        tableRows[tableCols[4]] = slDocs.GetCellValueAsString(i, iStartColumnIndex + 4);
                        tableRows[tableCols[5]] = slDocs.GetCellValueAsString(i, iStartColumnIndex + 5);
                        tableRows[tableCols[6]] = slDocs.GetCellValueAsString(i, iStartColumnIndex + 6);
                        tableRows[tableCols[7]] = slDocs.GetCellValueAsString(i, iStartColumnIndex + 7);

                        resultTable.Rows.Add(tableRows);
                        scrollToEOL();
                    }
                    
                }

            }
            else
            {
                Console.WriteLine("No File @ {0}.xlsx!", fullFilePath);
            }
        }

        #endregion

        #region ShiftControl
        private void toggleSwitch_ShiftAutoMan_Toggled(object sender, EventArgs e)
        {
            if (toggleSwitch_ShiftAutoMan.IsOn)
            {
                formInfo.autoShiftState = true;
            }
            else
            {
                formInfo.autoShiftState = false;
                if(cbb_ModelList.Items.Count > 0)
                {
                    cbb_ModelList.SelectedIndex = 0;
                }
            }
        }

        private void secTiming_Tick(object sender, EventArgs e)
        {
            String StrTimeNow_Hour = DateTime.Now.ToString("HH");
            String StrTimeNow_Min = DateTime.Now.ToString("mm");

            int timeNowHour = Convert.ToInt16(StrTimeNow_Hour);
            int timeNowMin = Convert.ToInt16(StrTimeNow_Min);

            //Console.WriteLine("{0}:{1} / {2}:{3} / {4}:{5}", timeNowHour.ToString(), timeNowMin.ToString(), Ca01_Hour,  Ca01_Min, Ca02_Hour, Ca02_Min);

            int totalTime = (timeNowHour * 60) + timeNowMin;
            int totalCa01 = (Ca01_Hour * 60) + Ca01_Min;
            int totalCa02 = (Ca02_Hour * 60) + Ca02_Min;

            syncContext.Post(
                new SendOrPostCallback(
                    delegate
                    {
                        if (this.toggleSwitch_ShiftAutoMan.IsOn)
                        {
                            if (totalTime.InRange(totalCa01, totalCa02))
                            {
                                if (this.cbb_Shift.Items.Count > 0)
                                {
                                    this.cbb_Shift.SelectedIndex = 0;
                                }
                            }
                            else
                            {
                                if (this.cbb_Shift.Items.Count > 1)
                                {
                                    this.cbb_Shift.SelectedIndex = 1;
                                }
                            }
                        }
                    }), null);

            
        }
        #endregion

        #region Save/Load Config
        private void saveConfig()
        {
            try
            {
                //SaveXML.SaveData(formInfo, "Config.xml");
                SaveXML.SaveData(formInfo, configFilePath + "\\" + configFileName);

                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void loadConfig()
        {
            //if (File.Exists("Config.xml"))
            if (File.Exists(configFilePath + "\\" + configFileName))
            {
                XmlSerializer xs = new XmlSerializer(typeof(Infomation));
                //FileStream read = new FileStream("Config.xml", FileMode.Open, FileAccess.Read, FileShare.Read);
                FileStream read = new FileStream(configFilePath + "\\" + configFileName, FileMode.Open, FileAccess.Read, FileShare.Read);
                formInfo = (Infomation)xs.Deserialize(read);
                read.Close();

                txtCamIP.Text = formInfo.CameraIP;
                txtCamUsr.Text = formInfo.CameraUsr;
                txtCamPwds.Text = formInfo.CameraPwds;

                
                checkAutoConnect.CheckState = formInfo.autoLogin ? CheckState.Checked : CheckState.Unchecked;
                chkbox_IsDMR.CheckState = formInfo.IsDMRCheckState ? CheckState.Checked : CheckState.Unchecked;
                chkbox_ShowHideDimension.CheckState = formInfo.showHideDimensionGrp ? CheckState.Checked : CheckState.Unchecked;

                toggleSwitch_ShiftAutoMan.IsOn = formInfo.autoShiftState;

                if (!String.IsNullOrEmpty(formInfo.PlcComport))
                {
                    if (cbbPlcPortList.Items.Contains(formInfo.PlcComport))
                    {
                        cbbPlcPortList.SelectedIndex = cbbPlcPortList.Items.IndexOf(formInfo.PlcComport);
                    }
                }


                if (!String.IsNullOrEmpty(formInfo.LogFileLocationPath))
                {
                    LogFilePath = formInfo.LogFileLocationPath;
                }
                
            }
        }
        #endregion

        #region Start/Stop System
        private void clickBtnConnectVision()
        {
            this.btnCamConnect.PerformClick();
        }

        private void clickBtnDisconnectVision()
        {
            this.btnCamDisconnect.PerformClick();
        }

        private void SystemStart()
        {
            
            syncContext.Post(
                new SendOrPostCallback(
                    delegate
                    {
                        this.cbb_Shift.Enabled = false;
                        this.cbb_ModelList.Enabled = false;
                        this.btn_Confirm.Enabled = false;

                        if (VisionConnectionState == (int)SensorState.Disconnected_)
                        {
                            btnCamConnect_Click(null, null);
                        }

                        if (SerialPortState == (int)PortState.Disconnected_)
                        {
                            btnDmrConnect_Click(null, null);
                        }

                        if (this.lblPlcStatus.Text == "Disconnected")
                        {
                            btnPlcConnect_Click(null, null);
                        }
                        
                    }), null);

            


            

            VMSystemState = true;
        }

        private void SystemStop()
        {
            this.cbb_Shift.Enabled = true;
            this.cbb_ModelList.Enabled = true;
            this.btn_Confirm.Enabled = false;
            VMSystemState = false;
            
            if (VisionConnectionState == (int)SensorState.Connected_)
            {
                btnCamDisconnect_Click(null, null);
            }


            if (SerialPortState == (int)PortState.Connected_)
            {
                btnDmrDisconnect_Click(null, null);
            }

            if (this.lblPlcStatus.Text == "Connected")
            {
                btnPlcDisconnect_Click(null, null);
            }
        }
        #endregion

        #region Display Alarm/Highlight

        private void displayFileVersion()
        {
            String currentName = this.Text;

            String currentVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            this.Text = currentName + " (Version " + currentVersion + ")";
        }

        private void setPlcTextStatus(string displayText, System.Drawing.Color displayColor)
        {
            lblPlcStatus.BackColor = displayColor;
            lblPlcStatus.Text = displayText;
        }

        //  Display on DMR Connecting
        private void displayDMRportState()
        {
            this.lblStatusHandheld.Text = SerialPortState_t[SerialPortState];

            switch (this.SerialPortState)
            {
                case (int)PortState.Connecting_:
                {
                    if (chkbox_IsDMR.CheckState == CheckState.Checked)
                    {
                        //  DataMan Devices

                        this.stt_handheld.Caption = "Connecting";
                        

                        this.cbbDmrList.Enabled = false;
                        this.cbbDmrBaud.Enabled = false;

                        this.chkbox_IsDMR.Enabled = false;

                        this.btnDmrConnect.Enabled = false;
                        this.btnDmrDisconnect.Enabled = false;


                        Console.WriteLine("[{0}] Connecting to Serial Port", DateTime.Now.ToString("H:mm:ss.fff"));
                    }
                    else if (chkbox_IsDMR.CheckState == CheckState.Unchecked)
                    {
                        //  Other Devices

                    }
                    break;
                }

                case (int)PortState.Connected_:
                {
                    if (chkbox_IsDMR.CheckState == CheckState.Checked)
                    {
                        //  DataMan Devices
                        
                        this.Observe.Enabled = true;    //  Start TMR to check connection

                        this.stt_handheld.Caption = "Connected";

                        this.cbbDmrList.Enabled = false;
                        this.cbbDmrBaud.Enabled = false;

                        this.chkbox_IsDMR.Enabled = false;

                        this.btnDmrConnect.Enabled = false;
                        this.btnDmrDisconnect.Enabled = true;

                        inCorruptSerialConnection = false;

                        Console.WriteLine("[{0}] Connected to Serial Port", DateTime.Now.ToString("H:mm:ss.fff"));
                    }
                    else if (chkbox_IsDMR.CheckState == CheckState.Unchecked)
                    {
                        //  Other Devices
                        this.stt_handheld.Caption = "Connected";

                        this.cbbDmrList.Enabled = false;
                        this.cbbDmrBaud.Enabled = false;

                        this.chkbox_IsDMR.Enabled = false;

                        this.btnDmrConnect.Enabled = false;
                        this.btnDmrDisconnect.Enabled = true;

                        //  Save config
                        formInfo.baudIndex = this.cbbDmrBaud.SelectedIndex;
                    }
                    break;
                }

                case (int)PortState.Disconnecting_:
                {
                    if (chkbox_IsDMR.CheckState == CheckState.Checked)
                    {
                        //  DataMan Devices

                    }
                    else if (chkbox_IsDMR.CheckState == CheckState.Unchecked)
                    {
                        //  Other Devices

                    }
                    break;
                }

                case (int)PortState.Disconnected_:
                {
                    if (chkbox_IsDMR.CheckState == CheckState.Checked)
                    {
                        //  DataMan Devices

                        

                        this.stt_handheld.Caption = "Disconnected";

                        this.cbbDmrList.Enabled = true;
                        this.cbbDmrBaud.Enabled = false;


                        this.btnDmrConnect.Enabled = true;
                        this.btnDmrDisconnect.Enabled = false;

                        this.chkbox_IsDMR.Enabled = true;
                        if (!inCorruptSerialConnection)
                        {
                            this.Observe.Enabled = false;
                        }
                    }
                    else if (chkbox_IsDMR.CheckState == CheckState.Unchecked)
                    {
                        //  Other Devices
                        this.stt_handheld.Caption = "Disconnected";

                        this.cbbDmrList.Enabled = true;
                        this.cbbDmrBaud.Enabled = true;

                        this.btnDmrConnect.Enabled = true;
                        this.btnDmrDisconnect.Enabled = false;

                        this.chkbox_IsDMR.Enabled = true;
                    }
                    break;
                }

                case (int)PortState.LostConnection_:
                {
                    if (chkbox_IsDMR.CheckState == CheckState.Checked)
                    {
                        //  DataMan Devices

                    }
                    else if (chkbox_IsDMR.CheckState == CheckState.Unchecked)
                    {
                        //  Other Devices

                    }
                    break;
                }

                case (int)PortState.ReConnecting_:
                {
                    if (chkbox_IsDMR.CheckState == CheckState.Checked)
                    {
                        //  DataMan Devices

                    }
                    else if (chkbox_IsDMR.CheckState == CheckState.Unchecked)
                    {
                        //  Other Devices

                    }
                    break;
                }
            }
        }

        //  Display on DMR Connected
        /*
        private void display_DmrConnected()
        {
            //  Addition
            this.Observe.Enabled = true;

            this.stt_handheld.Caption = "Connecting";

            this.cbbDmrList.Enabled = false;
            this.cbbDmrBaud.Enabled = false;

            this.chkbox_IsDMR.Enabled = false;

            this.btnDmrConnect.Enabled = false;
            this.btnDmrDisconnect.Enabled = true;

            //  Save config
            formInfo.baudIndex = this.cbbDmrBaud.SelectedIndex;

            //Console.WriteLine("Baud Index {0} / {1}", formInfo.baudIndex, cbbDmrBaud.SelectedIndex);
        }
        */
        //  Display on DMR cannot connect to target or disconnected
        /*
        private void display_DmrNotConnect()
        {
            //this.Observe.Enabled = false;


            //  Addition
            this.stt_handheld.Caption = "Disconnected";

            //  Serial Port setting
            if (chkbox_IsDMR.CheckState == CheckState.Checked)
            {
                this.cbbDmrList.Enabled = true;
                this.cbbDmrBaud.Enabled = false;
            }
            else if (chkbox_IsDMR.CheckState == CheckState.Unchecked)
            {
                this.cbbDmrList.Enabled = true;
                this.cbbDmrBaud.Enabled = true;
            }
            

            this.btnDmrConnect.Enabled = true;
            this.btnDmrDisconnect.Enabled = false;

            this.chkbox_IsDMR.Enabled = true;
        }
        */


        private void displayVisionConnectionState()
        {
            lblStatusCamera.Text = SensorState_t[VisionConnectionState];
            
            switch (VisionConnectionState)
            {
                case (int)SensorState.Disconnected_:
                case (int)SensorState.NotConnect_:
                {
                    //  Addition
                    this.btnCamConnect.Enabled = true;
                    this.btnCamDisconnect.Enabled = false;

                    this.txtCamIP.Enabled = true;
                    this.txtCamUsr.Enabled = true;
                    this.txtCamPwds.Enabled = true;

                    this.stt_CameraStatus.Caption = "Not Connect";

                    this.chkbox_ShowHideDimension.Enabled = false;

                    //  display Camera status
                    this.lblStatusCamera.Text = SensorState_t[VisionConnectionState];
                    this.lblStatusCamera.BackColor = System.Drawing.Color.Red;
                    break;
                }

                case (int)SensorState.Connecting_:
                {
                    this.btnCamConnect.Enabled = false;
                    this.btnCamDisconnect.Enabled = false;

                    this.txtCamIP.Enabled = false;
                    this.txtCamUsr.Enabled = false;
                    this.txtCamPwds.Enabled = false;

                    this.stt_CameraStatus.Caption = "Connecting";

                    //  display Camera status
                    this.lblStatusCamera.Text = SensorState_t[VisionConnectionState];
                    this.lblStatusCamera.BackColor = System.Drawing.Color.Orange;
                    break;
                }

                case (int)SensorState.Connected_:
                {
                    if (fullAccess)
                    {
                        //grpAddModel.Enabled = false;
                    }
                    else
                    {
                        //grpAddModel.Enabled = false;
                    }

                    //  Addition
                    this.btnCamConnect.Enabled = false;
                    this.btnCamDisconnect.Enabled = true;

                    this.txtCamIP.Enabled = false;
                    this.txtCamUsr.Enabled = false;
                    this.txtCamPwds.Enabled = false;

                    this.stt_CameraStatus.Caption = "Connected";

                    this.chkbox_ShowHideDimension.Enabled = true;

                    //  display Camera status
                    this.lblStatusCamera.Text = SensorState_t[VisionConnectionState];
                    this.lblStatusCamera.BackColor = System.Drawing.Color.Lime;
                    break;
                }

                case (int)SensorState.Offline_:
                {
                    this.lblStatusCamera.Text = SensorState_t[VisionConnectionState];
                    this.lblStatusCamera.BackColor = System.Drawing.Color.Orange;
                    break;
                }

                case (int)SensorState.Online_:
                {
                    this.lblStatusCamera.Text = SensorState_t[VisionConnectionState];
                    this.lblStatusCamera.BackColor = System.Drawing.Color.Lime;
                    break;
                }
            }

        }

        //  Display on Vision Connecting or Connected
        /*
        private void display_VisionConnected()
        {
            //  Addition
            this.btnCamConnect.Enabled = false;
            this.stt_CameraStatus.Caption = "Connecting...";
        }
        */

        //  Display on Vision cannot connect to target or disconnected
        /*
        private void display_VisionNotConnect()
        {
            //  Addition
            this.btnCamConnect.Enabled = true;
            this.btnCamDisconnect.Enabled = false;

            this.txtCamIP.Enabled = true;
            this.txtCamUsr.Enabled = true;
            this.txtCamPwds.Enabled = true;

            this.stt_CameraStatus.Caption = "Disconnected";
        }
        */

        //  Display Vision OCR text to Screen
        private void setVisionText(String visionString)
        {
            if (visionString.Length != Fixed_visionLen)
            {

            }
            else
            {
                char[] visionArr = new char[17];
                visionArr = visionString.ToCharArray();

                lblV01.Text = visionArr[0].ToString();
                lblV02.Text = visionArr[1].ToString();
                lblV03.Text = visionArr[2].ToString();
                lblV04.Text = visionArr[3].ToString();
                lblV05.Text = visionArr[4].ToString();
                lblV06.Text = visionArr[5].ToString();
                lblV07.Text = visionArr[6].ToString();
                lblV08.Text = visionArr[7].ToString();
                lblV09.Text = visionArr[8].ToString();
                lblV10.Text = visionArr[9].ToString();
                lblV11.Text = visionArr[10].ToString();
                lblV12.Text = visionArr[11].ToString();
                lblV13.Text = visionArr[12].ToString();
                lblV14.Text = visionArr[13].ToString();
                lblV15.Text = visionArr[14].ToString();
                lblV16.Text = visionArr[15].ToString();
                lblV17.Text = visionArr[16].ToString();

            }
        }

        private void setVisionText_v2(String visionString)
        {
            short tmpVisionLen = Fixed_visionLen;
            Fixed_visionLen = 13;
            if (visionString.Length != Fixed_visionLen)
            {

            }
            else
            {
                char[] visionArr = new char[13];
                visionArr = visionString.ToCharArray();

                lblV01.Text = visionArr[0].ToString();
                lblV02.Text = visionArr[1].ToString();
                lblV03.Text = visionArr[2].ToString();
                lblV04.Text = visionArr[3].ToString();
                lblV05.Text = visionArr[4].ToString();
                lblV06.Text = visionArr[5].ToString();
                lblV07.Text = visionArr[6].ToString();
                lblV08.Text = visionArr[7].ToString();
                lblV09.Text = visionArr[8].ToString();
                lblV10.Text = visionArr[9].ToString();
                lblV11.Text = visionArr[10].ToString();
                lblV12.Text = visionArr[11].ToString();
                lblV13.Text = visionArr[12].ToString();

            }

            Fixed_visionLen = tmpVisionLen;
        }

        //  Display Barcode date text to Screen
        private void setBarcodeText(String barcodeString)
        {
            Console.WriteLine("[{0}] Code length [{1}]: {2}", DateTime.Now.ToString("H:mm:ss.fff"), barcodeString.Length, barcodeString);

            if (barcodeString.Length != Fixed_barcodeLen)
            {

            }
            else
            {
                char[] barcodeArr = new char[17];
                barcodeArr = barcodeString.ToCharArray();
                displayHandheldString = barcodeString;

                lblB01.Text = barcodeArr[0].ToString();
                lblB02.Text = barcodeArr[1].ToString();
                lblB03.Text = barcodeArr[2].ToString();
                lblB04.Text = barcodeArr[3].ToString();
                lblB05.Text = barcodeArr[4].ToString();
                lblB06.Text = barcodeArr[5].ToString();
                lblB07.Text = barcodeArr[6].ToString();
                lblB08.Text = barcodeArr[7].ToString();
                lblB09.Text = barcodeArr[8].ToString();
                lblB10.Text = barcodeArr[9].ToString();
                lblB11.Text = barcodeArr[10].ToString();
                lblB12.Text = barcodeArr[11].ToString();
                lblB13.Text = barcodeArr[12].ToString();
                lblB14.Text = barcodeArr[13].ToString();
                lblB15.Text = barcodeArr[14].ToString();
                lblB16.Text = barcodeArr[15].ToString();
                lblB17.Text = barcodeArr[16].ToString();


            }

        }

        private void setBarcodeText_v2(String barcodeString)
        {
            Console.WriteLine("[{0}] Code length [{1}]: {2}", DateTime.Now.ToString("H:mm:ss.fff"), barcodeString.Length, barcodeString);

            short tmpBarcodeLen = Fixed_barcodeLen;
            Fixed_barcodeLen = 13;
            if (barcodeString.Length != Fixed_barcodeLen)
            {

            }
            else
            {
                char[] barcodeArr = new char[13];
                barcodeArr = barcodeString.ToCharArray();
                displayHandheldString = barcodeString;

                lblB01.Text = barcodeArr[0].ToString();
                lblB02.Text = barcodeArr[1].ToString();
                lblB03.Text = barcodeArr[2].ToString();
                lblB04.Text = barcodeArr[3].ToString();
                lblB05.Text = barcodeArr[4].ToString();
                lblB06.Text = barcodeArr[5].ToString();
                lblB07.Text = barcodeArr[6].ToString();
                lblB08.Text = barcodeArr[7].ToString();
                lblB09.Text = barcodeArr[8].ToString();
                lblB10.Text = barcodeArr[9].ToString();
                lblB11.Text = barcodeArr[10].ToString();
                lblB12.Text = barcodeArr[11].ToString();
                lblB13.Text = barcodeArr[12].ToString();
            }
            Fixed_barcodeLen = tmpBarcodeLen;
        }

        //  Error Blink Timer
        private void ErrBlinker_Tick(object sender, EventArgs e)
        {
            syncContext.Post(
                new SendOrPostCallback(
                    delegate
                    {
                        //  Button START Blink

                        if (this.btnStartSystem.Enabled == true)
                        {
                            if (this.btnStartSystem.BackColor == System.Drawing.Color.FromArgb(235, 236, 239))
                            {
                                this.btnStartSystem.BackColor = System.Drawing.Color.GreenYellow;
                            }
                            else
                            {
                                this.btnStartSystem.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
                            }
                        }
                        

                        //  Display Hand-held Status
                        switch (SerialPortState)
                        {
                            case (int)PortState.Connecting_:
                            {
                                this.lblStatusHandheld.Text = SerialPortState_t[SerialPortState];
                                this.lblStatusHandheld.BackColor = System.Drawing.Color.Lime;
                                break;
                            }
                            
                            case (int)PortState.Connected_:
                            {
                                this.lblStatusHandheld.Text = SerialPortState_t[SerialPortState];
                                this.lblStatusHandheld.BackColor = System.Drawing.Color.Lime;
                                break;
                            }

                            case (int)PortState.LostConnection_:
                            {
                                this.lblStatusHandheld.Text = SerialPortState_t[SerialPortState];
                                this.lblStatusHandheld.BackColor = System.Drawing.Color.Red;
                                break;
                            }

                            case (int)PortState.ReConnecting_:
                            {
                                this.lblStatusHandheld.Text = SerialPortState_t[SerialPortState];
                                this.lblStatusHandheld.BackColor = System.Drawing.Color.Orange;
                                break;
                            }

                            case (int)PortState.Disconnecting_:
                            {
                                this.lblStatusHandheld.Text = SerialPortState_t[SerialPortState];
                                this.lblStatusHandheld.BackColor = System.Drawing.Color.Red;
                                break;
                            }

                            case (int)PortState.Disconnected_:
                            {
                                this.lblStatusHandheld.Text = SerialPortState_t[SerialPortState];
                                this.lblStatusHandheld.BackColor = System.Drawing.Color.Red;
                                break;
                            }
                        }

                        //  Display Camera Status
                        switch (VisionConnectionState)
                        {
                            case (int)PortState.Connecting_:
                            {

                                break;
                            }

                            case (int)PortState.Connected_:
                            {
                                
                                break;
                            }

                            case (int)PortState.Disconnected_:
                            {
                                
                                break;
                            }

                            case (int)SensorState.NotConnect_:
                            {

                                break;
                            }

                            case (int)SensorState.Offline_:
                            {

                                break;
                            }

                            case (int)SensorState.Online_:
                            {

                                break;
                            }

                        }

                    }), null);
            
        }

        //  Display Result OK on screen
        private void displayResultOK()
        {
            this.lblResult_OK_NG.Text = "OK";
            this.lblResult_OK_NG.BackColor = System.Drawing.Color.Lime;
        }

        //  Display Result NG on screen
        private void displayResultNG()
        {
            this.lblResult_OK_NG.Text = "NG";
            this.lblResult_OK_NG.BackColor = System.Drawing.Color.Red;
        }

        //  Clear OK/NG Result to transparent
        private void clearResultOK_NG()
        {
            this.lblResult_OK_NG.Text = "";
            this.lblResult_OK_NG.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
        }

        //  Display Dimension on Screen
        private void displayCheckParameter(CvsCellCollection inputCells)
        {
           

            for (int i = 0; i < 5; i++)
            {
               if ((inputCells.GetCell(181, i + 2).Text == "1.000") || (inputCells.GetCell(181, i + 2).Text == "1,000"))
               {
                   checkList[i] = true;
               }
               else
               {
                   checkList[i] = false;
               }

               Console.WriteLine("Check List[{0}]: {1}", i, checkList[i].ToString());
            }


            if (checkList[0])
            {
                //  Height checking is enable
                //lblHeight.BackColor = System.Drawing.Color.Orange;
                lblHeight.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);

                lbl_H01.BorderStyle = BorderStyle.FixedSingle;
                lbl_H02.BorderStyle = BorderStyle.FixedSingle;
                lbl_H03.BorderStyle = BorderStyle.FixedSingle;
                lbl_H04.BorderStyle = BorderStyle.FixedSingle;
                lbl_H05.BorderStyle = BorderStyle.FixedSingle;
                lbl_H06.BorderStyle = BorderStyle.FixedSingle;
                lbl_H07.BorderStyle = BorderStyle.FixedSingle;
                lbl_H08.BorderStyle = BorderStyle.FixedSingle;
                lbl_H09.BorderStyle = BorderStyle.FixedSingle;
                lbl_H10.BorderStyle = BorderStyle.FixedSingle;
                lbl_H11.BorderStyle = BorderStyle.FixedSingle;
                lbl_H12.BorderStyle = BorderStyle.FixedSingle;
                lbl_H13.BorderStyle = BorderStyle.FixedSingle;
                lbl_H14.BorderStyle = BorderStyle.FixedSingle;
                lbl_H15.BorderStyle = BorderStyle.FixedSingle;
                lbl_H16.BorderStyle = BorderStyle.FixedSingle;
                lbl_H17.BorderStyle = BorderStyle.FixedSingle;

                lbl_H01.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_H02.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_H03.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_H04.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_H05.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_H06.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_H07.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_H08.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_H09.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_H10.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_H11.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_H12.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_H13.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_H14.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_H15.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_H16.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_H17.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
            }
            else
            {
                lblHeight.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);

                lbl_H01.BorderStyle = BorderStyle.None;
                lbl_H02.BorderStyle = BorderStyle.None;
                lbl_H03.BorderStyle = BorderStyle.None;
                lbl_H04.BorderStyle = BorderStyle.None;
                lbl_H05.BorderStyle = BorderStyle.None;
                lbl_H06.BorderStyle = BorderStyle.None;
                lbl_H07.BorderStyle = BorderStyle.None;
                lbl_H08.BorderStyle = BorderStyle.None;
                lbl_H09.BorderStyle = BorderStyle.None;
                lbl_H10.BorderStyle = BorderStyle.None;
                lbl_H11.BorderStyle = BorderStyle.None;
                lbl_H12.BorderStyle = BorderStyle.None;
                lbl_H13.BorderStyle = BorderStyle.None;
                lbl_H14.BorderStyle = BorderStyle.None;
                lbl_H15.BorderStyle = BorderStyle.None;
                lbl_H16.BorderStyle = BorderStyle.None;
                lbl_H17.BorderStyle = BorderStyle.None;

                lbl_H01.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_H02.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_H03.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_H04.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_H05.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_H06.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_H07.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_H08.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_H09.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_H10.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_H11.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_H12.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_H13.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_H14.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_H15.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_H16.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_H17.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
            }

            if (checkList[1])
            {
                //  Width checking is enable
                lblWidth.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);

                lbl_W01.BorderStyle = BorderStyle.FixedSingle;
                lbl_W02.BorderStyle = BorderStyle.FixedSingle;
                lbl_W03.BorderStyle = BorderStyle.FixedSingle;
                lbl_W04.BorderStyle = BorderStyle.FixedSingle;
                lbl_W05.BorderStyle = BorderStyle.FixedSingle;
                lbl_W06.BorderStyle = BorderStyle.FixedSingle;
                lbl_W07.BorderStyle = BorderStyle.FixedSingle;
                lbl_W08.BorderStyle = BorderStyle.FixedSingle;
                lbl_W09.BorderStyle = BorderStyle.FixedSingle;
                lbl_W10.BorderStyle = BorderStyle.FixedSingle;
                lbl_W11.BorderStyle = BorderStyle.FixedSingle;
                lbl_W12.BorderStyle = BorderStyle.FixedSingle;
                lbl_W13.BorderStyle = BorderStyle.FixedSingle;
                lbl_W14.BorderStyle = BorderStyle.FixedSingle;
                lbl_W15.BorderStyle = BorderStyle.FixedSingle;
                lbl_W16.BorderStyle = BorderStyle.FixedSingle;
                lbl_W17.BorderStyle = BorderStyle.FixedSingle;

                lbl_W01.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_W02.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_W03.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_W04.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_W05.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_W06.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_W07.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_W08.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_W09.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_W10.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_W11.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_W12.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_W13.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_W14.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_W15.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_W16.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_W17.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);

            }
            else
            {
                lblWidth.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);

                lbl_W01.BorderStyle = BorderStyle.None;
                lbl_W02.BorderStyle = BorderStyle.None;
                lbl_W03.BorderStyle = BorderStyle.None;
                lbl_W04.BorderStyle = BorderStyle.None;
                lbl_W05.BorderStyle = BorderStyle.None;
                lbl_W06.BorderStyle = BorderStyle.None;
                lbl_W07.BorderStyle = BorderStyle.None;
                lbl_W08.BorderStyle = BorderStyle.None;
                lbl_W09.BorderStyle = BorderStyle.None;
                lbl_W10.BorderStyle = BorderStyle.None;
                lbl_W11.BorderStyle = BorderStyle.None;
                lbl_W12.BorderStyle = BorderStyle.None;
                lbl_W13.BorderStyle = BorderStyle.None;
                lbl_W14.BorderStyle = BorderStyle.None;
                lbl_W15.BorderStyle = BorderStyle.None;
                lbl_W16.BorderStyle = BorderStyle.None;
                lbl_W17.BorderStyle = BorderStyle.None;

                lbl_W01.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_W02.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_W03.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_W04.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_W05.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_W06.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_W07.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_W08.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_W09.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_W10.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_W11.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_W12.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_W13.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_W14.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_W15.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_W16.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_W17.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
            }

            if (checkList[2])
            {
                lblAngle.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);

                lbl_A01.BorderStyle = BorderStyle.FixedSingle;
                lbl_A02.BorderStyle = BorderStyle.FixedSingle;
                lbl_A03.BorderStyle = BorderStyle.FixedSingle;
                lbl_A04.BorderStyle = BorderStyle.FixedSingle;
                lbl_A05.BorderStyle = BorderStyle.FixedSingle;
                lbl_A06.BorderStyle = BorderStyle.FixedSingle;
                lbl_A07.BorderStyle = BorderStyle.FixedSingle;
                lbl_A08.BorderStyle = BorderStyle.FixedSingle;
                lbl_A09.BorderStyle = BorderStyle.FixedSingle;
                lbl_A10.BorderStyle = BorderStyle.FixedSingle;
                lbl_A11.BorderStyle = BorderStyle.FixedSingle;
                lbl_A12.BorderStyle = BorderStyle.FixedSingle;
                lbl_A13.BorderStyle = BorderStyle.FixedSingle;
                lbl_A14.BorderStyle = BorderStyle.FixedSingle;
                lbl_A15.BorderStyle = BorderStyle.FixedSingle;
                lbl_A16.BorderStyle = BorderStyle.FixedSingle;
                lbl_A17.BorderStyle = BorderStyle.FixedSingle;

                lbl_A01.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_A02.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_A03.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_A04.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_A05.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_A06.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_A07.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_A08.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_A09.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_A10.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_A11.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_A12.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_A13.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_A14.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_A15.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_A16.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_A17.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
            }
            else
            {
                lblAngle.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);

                lbl_A01.BorderStyle = BorderStyle.None;
                lbl_A02.BorderStyle = BorderStyle.None;
                lbl_A03.BorderStyle = BorderStyle.None;
                lbl_A04.BorderStyle = BorderStyle.None;
                lbl_A05.BorderStyle = BorderStyle.None;
                lbl_A06.BorderStyle = BorderStyle.None;
                lbl_A07.BorderStyle = BorderStyle.None;
                lbl_A08.BorderStyle = BorderStyle.None;
                lbl_A09.BorderStyle = BorderStyle.None;
                lbl_A10.BorderStyle = BorderStyle.None;
                lbl_A11.BorderStyle = BorderStyle.None;
                lbl_A12.BorderStyle = BorderStyle.None;
                lbl_A13.BorderStyle = BorderStyle.None;
                lbl_A14.BorderStyle = BorderStyle.None;
                lbl_A15.BorderStyle = BorderStyle.None;
                lbl_A16.BorderStyle = BorderStyle.None;
                lbl_A17.BorderStyle = BorderStyle.None;

                lbl_A01.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_A02.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_A03.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_A04.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_A05.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_A06.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_A07.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_A08.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_A09.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_A10.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_A11.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_A12.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_A13.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_A14.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_A15.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_A16.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_A17.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
            }


            if (checkList[3])
            {
                lblCenterDistance.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);

                lbl_C01.BorderStyle = BorderStyle.FixedSingle;
                lbl_C02.BorderStyle = BorderStyle.FixedSingle;
                lbl_C03.BorderStyle = BorderStyle.FixedSingle;
                lbl_C04.BorderStyle = BorderStyle.FixedSingle;
                lbl_C05.BorderStyle = BorderStyle.FixedSingle;
                lbl_C06.BorderStyle = BorderStyle.FixedSingle;
                lbl_C07.BorderStyle = BorderStyle.FixedSingle;
                lbl_C08.BorderStyle = BorderStyle.FixedSingle;
                lbl_C09.BorderStyle = BorderStyle.FixedSingle;
                lbl_C10.BorderStyle = BorderStyle.FixedSingle;
                lbl_C11.BorderStyle = BorderStyle.FixedSingle;
                lbl_C12.BorderStyle = BorderStyle.FixedSingle;
                lbl_C13.BorderStyle = BorderStyle.FixedSingle;
                lbl_C14.BorderStyle = BorderStyle.FixedSingle;
                lbl_C15.BorderStyle = BorderStyle.FixedSingle;
                lbl_C16.BorderStyle = BorderStyle.FixedSingle;

                lbl_C01.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_C02.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_C03.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_C04.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_C05.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_C06.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_C07.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_C08.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_C09.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_C10.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_C11.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_C12.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_C13.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_C14.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_C15.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_C16.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
            }
            else
            {
                lblCenterDistance.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);

                lbl_C01.BorderStyle = BorderStyle.None;
                lbl_C02.BorderStyle = BorderStyle.None;
                lbl_C03.BorderStyle = BorderStyle.None;
                lbl_C04.BorderStyle = BorderStyle.None;
                lbl_C05.BorderStyle = BorderStyle.None;
                lbl_C06.BorderStyle = BorderStyle.None;
                lbl_C07.BorderStyle = BorderStyle.None;
                lbl_C08.BorderStyle = BorderStyle.None;
                lbl_C09.BorderStyle = BorderStyle.None;
                lbl_C10.BorderStyle = BorderStyle.None;
                lbl_C11.BorderStyle = BorderStyle.None;
                lbl_C12.BorderStyle = BorderStyle.None;
                lbl_C13.BorderStyle = BorderStyle.None;
                lbl_C14.BorderStyle = BorderStyle.None;
                lbl_C15.BorderStyle = BorderStyle.None;
                lbl_C16.BorderStyle = BorderStyle.None;

                lbl_C01.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_C02.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_C03.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_C04.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_C05.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_C06.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_C07.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_C08.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_C09.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_C10.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_C11.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_C12.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_C13.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_C14.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_C15.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_C16.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
            }


            if (checkList[4])
            {
                lblLineAngle.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
                lbl_LA01.BorderStyle = BorderStyle.FixedSingle;
                lbl_LA01.ForeColor = System.Drawing.Color.FromArgb(32, 31, 53);
            }
            else
            {
                lblLineAngle.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
                lbl_LA01.BorderStyle = BorderStyle.None;
                lbl_LA01.ForeColor = System.Drawing.Color.FromArgb(235, 236, 239);
            }

        }

        //  Clear all Labels except OK/NG
        private void clearResultLabel(bool isClearAll)
        {
            if (isClearAll)
            {
                displayHandheldString = "";
                lblB01.Text = "";
                lblB02.Text = "";
                lblB03.Text = "";
                lblB04.Text = "";
                lblB05.Text = "";
                lblB06.Text = "";
                lblB07.Text = "";
                lblB08.Text = "";
                lblB09.Text = "";
                lblB10.Text = "";
                lblB11.Text = "";
                lblB12.Text = "";
                lblB13.Text = "";
                lblB14.Text = "";
                lblB15.Text = "";
                lblB16.Text = "";
                lblB17.Text = "";
            }

            lblV01.Text = "";
            lblV02.Text = "";
            lblV03.Text = "";
            lblV04.Text = "";
            lblV05.Text = "";
            lblV06.Text = "";
            lblV07.Text = "";
            lblV08.Text = "";
            lblV09.Text = "";
            lblV10.Text = "";
            lblV11.Text = "";
            lblV12.Text = "";
            lblV13.Text = "";
            lblV14.Text = "";
            lblV15.Text = "";
            lblV16.Text = "";
            lblV17.Text = "";

            lbl_H01.Text = "";
            lbl_H02.Text = "";
            lbl_H03.Text = "";
            lbl_H04.Text = "";
            lbl_H05.Text = "";
            lbl_H06.Text = "";
            lbl_H07.Text = "";
            lbl_H08.Text = "";
            lbl_H09.Text = "";
            lbl_H10.Text = "";
            lbl_H11.Text = "";
            lbl_H12.Text = "";
            lbl_H13.Text = "";
            lbl_H14.Text = "";
            lbl_H15.Text = "";
            lbl_H16.Text = "";
            lbl_H17.Text = "";

            lbl_W01.Text = "";
            lbl_W02.Text = "";
            lbl_W03.Text = "";
            lbl_W04.Text = "";
            lbl_W05.Text = "";
            lbl_W06.Text = "";
            lbl_W07.Text = "";
            lbl_W08.Text = "";
            lbl_W09.Text = "";
            lbl_W10.Text = "";
            lbl_W11.Text = "";
            lbl_W12.Text = "";
            lbl_W13.Text = "";
            lbl_W14.Text = "";
            lbl_W15.Text = "";
            lbl_W16.Text = "";
            lbl_W17.Text = "";

            lbl_A01.Text = "";
            lbl_A02.Text = "";
            lbl_A03.Text = "";
            lbl_A04.Text = "";
            lbl_A05.Text = "";
            lbl_A06.Text = "";
            lbl_A07.Text = "";
            lbl_A08.Text = "";
            lbl_A09.Text = "";
            lbl_A10.Text = "";
            lbl_A11.Text = "";
            lbl_A12.Text = "";
            lbl_A13.Text = "";
            lbl_A14.Text = "";
            lbl_A15.Text = "";
            lbl_A16.Text = "";
            lbl_A17.Text = "";

            lbl_C01.Text = "";
            lbl_C02.Text = "";
            lbl_C03.Text = "";
            lbl_C04.Text = "";
            lbl_C05.Text = "";
            lbl_C06.Text = "";
            lbl_C07.Text = "";
            lbl_C08.Text = "";
            lbl_C09.Text = "";
            lbl_C10.Text = "";
            lbl_C11.Text = "";
            lbl_C12.Text = "";
            lbl_C13.Text = "";
            lbl_C14.Text = "";
            lbl_C15.Text = "";
            lbl_C16.Text = "";

            lbl_LA01.Text = "";
        }

        // Display Compare results between Barcode Data & Vision Data
        private void displayCompareResult()
        {
            //  Compare char at 01
            if (lblB01.Text == lblV01.Text)
            {
                lblB01.BackColor = System.Drawing.Color.LimeGreen;
                lblV01.BackColor = System.Drawing.Color.LimeGreen;
            }
            else
            {
                lblB01.BackColor = System.Drawing.Color.Red;
                lblV01.BackColor = System.Drawing.Color.Red;
            }

            //  Compare char at 02
            if (lblB02.Text == lblV02.Text)
            {
                lblB02.BackColor = System.Drawing.Color.LimeGreen;
                lblV02.BackColor = System.Drawing.Color.LimeGreen;
            }
            else
            {
                lblB02.BackColor = System.Drawing.Color.Red;
                lblV02.BackColor = System.Drawing.Color.Red;
            }

            //  Compare char at 03
            if (lblB03.Text == lblV03.Text)
            {
                lblB03.BackColor = System.Drawing.Color.LimeGreen;
                lblV03.BackColor = System.Drawing.Color.LimeGreen;
            }
            else
            {
                lblB03.BackColor = System.Drawing.Color.Red;
                lblV03.BackColor = System.Drawing.Color.Red;
            }

            //  Compare char at 04
            if (lblB04.Text == lblV04.Text)
            {
                lblB04.BackColor = System.Drawing.Color.LimeGreen;
                lblV04.BackColor = System.Drawing.Color.LimeGreen;
            }
            else
            {
                lblB04.BackColor = System.Drawing.Color.Red;
                lblV04.BackColor = System.Drawing.Color.Red;
            }

            //  Compare char at 05
            if (lblB05.Text == lblV05.Text)
            {
                lblB05.BackColor = System.Drawing.Color.LimeGreen;
                lblV05.BackColor = System.Drawing.Color.LimeGreen;
            }
            else
            {
                lblB05.BackColor = System.Drawing.Color.Red;
                lblV05.BackColor = System.Drawing.Color.Red;
            }

            //  Compare char at 06
            if (lblB06.Text == lblV06.Text)
            {
                lblB06.BackColor = System.Drawing.Color.LimeGreen;
                lblV06.BackColor = System.Drawing.Color.LimeGreen;
            }
            else
            {
                lblB06.BackColor = System.Drawing.Color.Red;
                lblV06.BackColor = System.Drawing.Color.Red;
            }

            //  Compare char at 07
            if (lblB07.Text == lblV07.Text)
            {
                lblB07.BackColor = System.Drawing.Color.LimeGreen;
                lblV07.BackColor = System.Drawing.Color.LimeGreen;
            }
            else
            {
                lblB07.BackColor = System.Drawing.Color.Red;
                lblV07.BackColor = System.Drawing.Color.Red;
            }

            //  Compare char at 08
            if (lblB08.Text == lblV08.Text)
            {
                lblB08.BackColor = System.Drawing.Color.LimeGreen;
                lblV08.BackColor = System.Drawing.Color.LimeGreen;
            }
            else
            {
                lblB08.BackColor = System.Drawing.Color.Red;
                lblV08.BackColor = System.Drawing.Color.Red;
            }

            //  Compare char at 09
            if (lblB09.Text == lblV09.Text)
            {
                lblB09.BackColor = System.Drawing.Color.LimeGreen;
                lblV09.BackColor = System.Drawing.Color.LimeGreen;
            }
            else
            {
                lblB09.BackColor = System.Drawing.Color.Red;
                lblV09.BackColor = System.Drawing.Color.Red;
            }

            //  Compare char at 10
            if (lblB10.Text == lblV10.Text)
            {
                lblB10.BackColor = System.Drawing.Color.LimeGreen;
                lblV10.BackColor = System.Drawing.Color.LimeGreen;
            }
            else
            {
                lblB10.BackColor = System.Drawing.Color.Red;
                lblV10.BackColor = System.Drawing.Color.Red;
            }

            //  Compare char at 11
            if (lblB11.Text == lblV11.Text)
            {
                lblB11.BackColor = System.Drawing.Color.LimeGreen;
                lblV11.BackColor = System.Drawing.Color.LimeGreen;
            }
            else
            {
                lblB11.BackColor = System.Drawing.Color.Red;
                lblV11.BackColor = System.Drawing.Color.Red;
            }

            //  Compare char at 12
            if (lblB12.Text == lblV12.Text)
            {
                lblB12.BackColor = System.Drawing.Color.LimeGreen;
                lblV12.BackColor = System.Drawing.Color.LimeGreen;
            }
            else
            {
                lblB12.BackColor = System.Drawing.Color.Red;
                lblV12.BackColor = System.Drawing.Color.Red;
            }

            //  Compare char at 13
            if (lblB13.Text == lblV13.Text)
            {
                lblB13.BackColor = System.Drawing.Color.LimeGreen;
                lblV13.BackColor = System.Drawing.Color.LimeGreen;
            }
            else
            {
                lblB13.BackColor = System.Drawing.Color.Red;
                lblV13.BackColor = System.Drawing.Color.Red;
            }

            //  Compare char at 14
            if (lblB14.Text == lblV14.Text)
            {
                lblB14.BackColor = System.Drawing.Color.LimeGreen;
                lblV14.BackColor = System.Drawing.Color.LimeGreen;
            }
            else
            {
                lblB14.BackColor = System.Drawing.Color.Red;
                lblV14.BackColor = System.Drawing.Color.Red;
            }

            //  Compare char at 15
            if (lblB15.Text == lblV15.Text)
            {
                lblB15.BackColor = System.Drawing.Color.LimeGreen;
                lblV15.BackColor = System.Drawing.Color.LimeGreen;
            }
            else
            {
                lblB15.BackColor = System.Drawing.Color.Red;
                lblV15.BackColor = System.Drawing.Color.Red;
            }

            //  Compare char at 16
            if (lblB16.Text == lblV16.Text)
            {
                lblB16.BackColor = System.Drawing.Color.LimeGreen;
                lblV16.BackColor = System.Drawing.Color.LimeGreen;
            }
            else
            {
                lblB16.BackColor = System.Drawing.Color.Red;
                lblV16.BackColor = System.Drawing.Color.Red;
            }

            //  Compare char at 17
            if (lblB17.Text == lblV17.Text)
            {
                lblB17.BackColor = System.Drawing.Color.LimeGreen;
                lblV17.BackColor = System.Drawing.Color.LimeGreen;
            }
            else
            {
                lblB17.BackColor = System.Drawing.Color.Red;
                lblV17.BackColor = System.Drawing.Color.Red;
            }
        }

        //  Clear all Results were displayed to transparent
        private void displayResultReset()
        {
            lblB01.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblB02.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblB03.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblB04.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblB05.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblB06.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblB07.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblB08.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblB09.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblB10.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblB11.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblB12.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblB13.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblB14.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblB15.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblB16.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblB17.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);

            lblV01.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblV02.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblV03.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblV04.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblV05.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblV06.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblV07.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblV08.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblV09.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblV10.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblV11.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblV12.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblV13.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblV14.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblV15.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblV16.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lblV17.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);

            lbl_H01.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_H02.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_H03.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_H04.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_H05.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_H06.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_H07.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_H08.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_H09.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_H10.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_H11.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_H12.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_H13.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_H14.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_H15.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_H16.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_H17.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);

            lbl_W01.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_W02.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_W03.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_W04.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_W05.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_W06.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_W07.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_W08.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_W09.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_W10.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_W11.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_W12.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_W13.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_W14.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_W15.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_W16.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_W17.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);

            lbl_A01.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_A02.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_A03.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_A04.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_A05.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_A06.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_A07.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_A08.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_A09.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_A10.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_A11.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_A12.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_A13.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_A14.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_A15.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_A16.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_A17.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);

            lbl_C01.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_C02.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_C03.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_C04.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_C05.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_C06.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_C07.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_C08.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_C09.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_C10.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_C11.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_C12.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_C13.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_C14.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_C15.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);
            lbl_C16.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);

            lbl_LA01.BackColor = System.Drawing.Color.FromArgb(235, 236, 239);

        }


        private void chkbox_ShowHideDimension_CheckedChanged(object sender, EventArgs e)
        {
            if (fullAccess)
            {
                if (chkbox_ShowHideDimension.CheckState == CheckState.Checked)
                {
                    formInfo.showHideDimensionGrp = true;
                    this.layoutControlItem4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                }
                else if (chkbox_ShowHideDimension.CheckState == CheckState.Unchecked)
                {
                    formInfo.showHideDimensionGrp = false;
                    this.layoutControlItem4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                }
            }

            
        }


        //  Show/Hide Compare results
        private void TMR_DisplayDelay_Tick(object sender, EventArgs e)
        {
            this.TMR_DisplayDelay.Enabled = false;
            isHandheldRead = true;
            

            //  Reset GUI Components to default color.
            displayResultReset();   //  Clear all format background
            clearResultOK_NG(); //  Clear OK/NG background & data
            clearResultLabel(false);    //  Clear all label except bar-code data

            if (VMSystemState)
            {
                Results_HandheldNow = Results_HandheldQueue[queuePosition];
                Results_HandheldQueue[queuePosition] = "";
                
                queueLock = false;

                isVisionRead = false;

                isQueueUp = false;

                Console.WriteLine("[{0}] TMR Display Queue[{1}]: {2}", DateTime.Now.ToString("H:mm:ss.fff"), queuePosition, Results_HandheldNow);

                //  Append to DataTable
                addRowData();
                addBarcodeData(Results_HandheldNow);

                

                //  Displays
                setBarcodeText_v2(Results_HandheldNow);   // display on screen
                setTransferStr(Fixed_barcodeLen, Results_HandheldNow); //  Send to Camera via Modbus TCP

                
            }
            
        }
        #endregion




        private void btnTest_Click(object sender, EventArgs e)
        {
            addRowData();
            addBarcodeData("BARCODE");
            addVisionData("VISION");
            addResult_OK_NG(true);
            saveSpreadsheet();
            syncContext.Post(
                delegate
                {

                }, null);
        }

        private void btnExportLocation_Click(object sender, EventArgs e)
        {
            //  Show the folderBrowserDialog.
            DialogResult results = folderBrowserDialog1.ShowDialog();

            if (results == DialogResult.OK)
            {
                LogFilePath = folderBrowserDialog1.SelectedPath;
                formInfo.LogFileLocationPath = LogFilePath;
            }

            
            syncContext.Post(
                delegate
                {

                }, null);
        }

        private void btnTrigger_Click(object sender, EventArgs e)
        {
            loadSpreadsheet();
        }


        private void LoadJogFileFromInSight()
        {
            if (inSight != null && this.VisionConnectionState == (int)SensorState.Connected_)
            {
                // Retrieve the file list from the sensor
                lbc_JobinCam.Items.Clear();
                mFileList = inSight.File.GetFileList();
                for (int i = 0; i < mFileList.Length; i++)
                {
                    // Filter out non-job files
                    if (Path.GetExtension(mFileList[i]).ToUpper() == ".JOB")
                    {
                        jobnamesincam.Add(mFileList[i]);
                        lbc_JobinCam.Items.Add(mFileList[i]);
                    }
                }
                inSight.LoadCompleted += new Cognex.InSight.CvsLoadCompletedEventHandler(mInSight_LoadCompleted);
            }
            else
            {
                Console.WriteLine("Không có kết nối với camera");
            }
        }

        private void mInSight_LoadCompleted(object sender, CvsLoadCompletedEventArgs e)
        {
            Console.WriteLine("Load Done");
        }

        private void OpenJobFileToInSight(string jobfile)
        {
            if (inSight != null)
            {
                bool state;

                // Save current state
                state = inSight.SoftOnline;

                // Must be offline to load a job
                inSight.SoftOnline = false;

                // Load the new job fie if it exits.
                if (jobfile != "")
                {
                    //freezeScreen();
                    inSight.File.LoadJobFile(jobfile);
                }

                // Restore the online state
                inSight.SoftOnline = state;
                Console.WriteLine("Load file job complete");
                //btnStartSystem_Click(null, null);
            }
            else
            {
                Console.WriteLine("Camera disconnect");
            }
        }

        private void lbc_ModelList_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtModelInput.Text = lbc_ModelList.GetItemText(lbc_ModelList.SelectedItem);
            lbc_JobList.SelectedIndex = lbc_ModelList.SelectedIndex;
        }

        private void lbc_JobList_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtJobInput.Text = lbc_JobList.GetItemText(lbc_JobList.SelectedItem);
            lbc_ModelList.SelectedIndex = lbc_JobList.SelectedIndex;
        }

        private void changeitems()
        {

        }

        private void lbc_JobinCam_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtJobInput.Text = lbc_JobinCam.GetItemText(lbc_JobinCam.SelectedItem);
        }

        private void cbb_ModelList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (jobnames.Count != 0 && cbb_ModelList.SelectedIndex != -1)
            {
                lbljobfile.Text = jobnames[cbb_ModelList.SelectedIndex];

            }
        }

        private void btn_Confirm_Click(object sender, EventArgs e)
        {

            //OpenJobFileToInSight(jobnames[cbb_ModelList.SelectedIndex]);
        }

        private void cbb_ModelList_SelectedValueChanged(object sender, EventArgs e)
        {
            //if (jobnames.Count != 0)
            //{
            //    lbljobfile.Text = jobnames[cbb_ModelList.SelectedIndex];

            //}
            //MessageBox.Show("ng");
        }

        private int PLCReadDevice(string address)
        {
            int resultDeviceValue = 0;

            try
            {
                int iReturnCode = this.axActProgType1.GetDevice(address, out resultDeviceValue);
                if (iReturnCode != 0)
                {
                    //  Error
                    Console.WriteLine("Error while reading 0x{0:X}", iReturnCode);
                }
                else
                {
                    return resultDeviceValue;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            return -1;
        }

        private void txtEdit_Click(object sender, EventArgs e)
        {
            index_select_before = -1;
            if (txtModelInput.Text != String.Empty)
            {
                modelnames[lbc_ModelList.SelectedIndex] = txtModelInput.Text;
                jobnames[lbc_ModelList.SelectedIndex] = txtJobInput.Text;
                lbc_JobList.DataSource = null;
                lbc_JobList.DataSource = jobnames;
                lbc_ModelList.DataSource = null;
                lbc_ModelList.DataSource = modelnames;
                cbb_ModelList.DataSource = null;
                cbb_ModelList.DataSource = modelnames;
            }
        }

        private void hideContainerLeft_Click(object sender, EventArgs e)
        {

        }
    }




















    public static class IComparableExtension
    {
        public static bool InRange<T>(this T value, T from, T to) where T : IComparable<T>
        {
            //return (value.CompareTo(from) >= 1) && (value.CompareTo(to) <= -1);   //  Does not accept equal from/to limit.
            return (value.CompareTo(from) >= 0) && (value.CompareTo(to) <= 0);      //  accept equal from/to limit.
        }
    }
}
