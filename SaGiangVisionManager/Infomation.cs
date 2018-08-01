using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;


namespace SaGiangVisionManager
{
    public class Infomation
    {

        private string CamIP;
        private string CamUsr;
        private string CamPwd;
        private bool autoLoginCheck = false;

        private string modelListString = "";
        private string jobListString = "";

        private string currentmodel = "";
        private bool AutoShift;

        //private bool shiftManagement;

        private bool showHideDimension = true;

        //private DataTable reserveTable;

        private string logFileLocation = "";
        
        //  Comport Setting
        private int comportbaudIndex = 0;
        private bool isDMRcheckState = false;

        //  Plc Port Setting

        private string _plcComport = "";

        // 

        public String CurrentModel
        {
            get { return currentmodel; }
            set { currentmodel = value; }
        }
        
        public String CameraIP
        {
            get{return CamIP;}
            set{CamIP = value;}
        }

        public String CameraUsr
        {
            get {return CamUsr;}
            set{CamUsr = value;}
        }

        public String CameraPwds
        {
            get { return CamPwd;}
            set { CamPwd = value;}
        }



        public bool autoLogin
        {
            get { return autoLoginCheck; }
            set { autoLoginCheck = value; }
        }

        //  Save/Load Listbox
        public string modelList
        {
            get { return modelListString; }
            set { modelListString = value; }
        }

        public string jobList
        {
            get { return jobListString; }
            set { jobListString = value; }
        }


        public bool autoShiftState
        {
            get { return AutoShift; }
            set { AutoShift = value; }
        }

        public String PlcComport
        {
            get { return _plcComport; }
            set { _plcComport = value; }
        }

//         public DataTable backupTable
//         {
//             get { return reserveTable; }
//             set { reserveTable = value; }
//         }

        public int baudIndex
        {
            get { return comportbaudIndex; }
            set { comportbaudIndex = value; }
        }

        public bool IsDMRCheckState
        {
            get { return isDMRcheckState; }
            set { isDMRcheckState = value; }
        }

        public bool showHideDimensionGrp
        {
            get { return showHideDimension; }
            set { showHideDimension = value; }
        }

        public String LogFileLocationPath
        {
            get { return logFileLocation; }
            set { logFileLocation = value; }
        }
    }
}
