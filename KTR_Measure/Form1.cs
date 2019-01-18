using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Net;               //載入網路
using System.Net.Sockets;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms.DataVisualization.Charting;

using PCI_DMC;
using PCI_DMC_ERR;


namespace KTR_Measure
{
    
    public partial class Form1 : Form
    {
        
        public const int rpmRate1 = 400; //ktr比例
        public const int rpmRate2 = 4;
        public const int torqueRate_10 = 1;
        public const int torqueRate_50 = 5;
        public const ushort node1 = 1; //節點    虎尾3.4 中山1.2

        Thread ThWorking_PLC;
        Socket T;
        
        int excelTime = 0;  //excel陣列數目

        short existcard = 0, rc;
        ushort gCardNo = 0, DeviceInfo = 0, gnodeid;
        ushort[] gCardNoList = new ushort[16];
        uint[] SlaveTable = new uint[4];
        ushort[] NodeID = new ushort[32];
        byte[] value = new byte[10];
        ushort gNodeNum;
        short spd1 = 0, toe1 = 0;
        bool ServoWorking = false;
        TextBox[] txtIoSts = new TextBox[16];


        List<double> ktrTorque1 = new List<double>();
        List<double> ktrTorque2 = new List<double>();
        List<double> ktrRpm1 = new List<double>();
        List<double> ktrRpm2 = new List<double>();
        List<double> motorTorque1 = new List<double>();
        List<double> motorRpm1 = new List<double>();
        double[,] rpm_1 = new double[90000, 10];
        double[,] rpm_2 = new double[90000, 10];
        double[] rpm_motor1 = new double[90000];
        double[] rpm_motor2 = new double[90000];
        double[,] torque_1 = new double[90000, 10];
        double[,] torque_2 = new double[90000, 10];
        double[] torque_motor1 = new double[90000];
        double[] torque_motor2 = new double[90000];

        public Form1()
        {
            InitializeComponent();
        }

        public void ArrayReset(double[] a)
        {
            for (int i = 0; i < a.Length; i++)
            {
                a[i] = 0;
            }
        }
        public void ArrayReset(double[,] a)
        {
            for (int i = 0; i < excelTime; i++)
                for (int j = 0; j < 10; j++)
                {
                    a[i, j] = 0;
                }
        }

        private void LogOutput(string s)
        {
            string Output = "";
            Output += DateTime.Now.ToString("HH:mm:ss>>>");
            Output += s;
            Output += "\n";
            dialog.AppendText(Output);
        }

        private void btnConnectPLC_Click(object sender, EventArgs e)
        {
            LogOutput("開始連線至PLC\n");
            string IP = txtIPToPLC.Text;                //設定變數IP，其字串
            int Port = int.Parse(txtPortToPLC.Text);    //設定變數Port，為整數
            try
            {
                IPEndPoint EP = new IPEndPoint(IPAddress.Parse(IP), Port);
                T = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                T.Connect(EP); //建立連線
                lblConnectStatus.Text = "已連線至PLC";
                LogOutput("已連線至PLC");
                ServoCon.Enabled = true;
                ThWorking_PLC = new Thread(working_PLC);
                ThWorking_PLC.Start();
            }
            catch (Exception)
            {
                lblConnectStatus.Text = "無法連線至PLC,請檢查線路或IP";
                LogOutput("無法連線至PLC,請檢查線路或IP");
                return;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;
            CType.SelectedIndex = 0;
            LogOutput("Welcome");
            LogOutput("Notice!!:");
            LogOutput("比例設定為：(rpm/Troque)");
            LogOutput("輸入端：(" + rpmRate1.ToString() + "/" + torqueRate_10.ToString() + ")");
            LogOutput("輸出端：(" + rpmRate2.ToString() + "/" + torqueRate_50.ToString() + ")");
        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Close();
            Environment.Exit(Environment.ExitCode);
        }

        private void btnSaveExcel_Click(object sender, EventArgs e)
        {
            //ThWorking_PLC.Abort();
            String FileStr = "D:\\";
            FileStr += DateTime.Now.ToString("yyyy-MM-dd_HHmmss");
            LogOutput("檔案儲存中.....");
            Excel.Application Excel_app1 = new Excel.Application();
            Excel.Workbook Excel_WB1 = Excel_app1.Workbooks.Add();
            Excel.Worksheet Excel_WS1 = new Excel.Worksheet();
           
            Excel_app1.Cells[1, 1] = "INPUT端轉速";
            Excel_app1.Cells[1, 2] = "INPUT端扭矩";
            Excel_app1.Cells[1, 3] = "OUTPUT端轉速";
            Excel_app1.Cells[1, 4] = "OUTPUT端扭矩";
            Excel_app1.Cells[1, 5] = "Motor轉速";
            Excel_app1.Cells[1, 6] = "Motor扭矩";

            for (int i = 0; i < ktrRpm1.Count; i++)
            {
                Excel_app1.Cells[i + 2, 1] = ktrRpm1[i];
                Excel_app1.Cells[i + 2, 2] = ktrTorque1[i];
                Excel_app1.Cells[i + 2, 3] = ktrRpm2[i];
                Excel_app1.Cells[i + 2, 4] = ktrTorque2[i];
                Excel_app1.Cells[i + 2, 5] = motorRpm1[i];
                Excel_app1.Cells[i + 2, 6] = motorTorque1[i];

            }
            Excel_WB1.SaveAs(FileStr);
            Excel_WB1.Close();
            Excel_WB1 = null;
            Excel_app1.Quit();
            Excel_app1 = null;
            LogOutput("檔案已儲存至：" + FileStr + ".xlsx");

            if (IfChart.Checked)
            {
                chart1.SaveImage(FileStr + "_INPUT_RPM", ChartImageFormat.Jpeg);
                chart2.SaveImage(FileStr + "_OUTPUT_RPM", ChartImageFormat.Jpeg);
                chart3.SaveImage(FileStr + "_INPUT_Torq", ChartImageFormat.Jpeg);
                chart4.SaveImage(FileStr + "_OUTPUT_Torq", ChartImageFormat.Jpeg);
                chart5.SaveImage(FileStr + "_MOTOR_RPM", ChartImageFormat.Jpeg);
                chart6.SaveImage(FileStr + "_MOTOR_Torq", ChartImageFormat.Jpeg);
                LogOutput("圖表輸出完成");
            }
        }

        private void btnreset1_Click(object sender, EventArgs e)
        {
            gnodeid = ushort.Parse(cmbNodeID.Text);
            CPCI_DMC.CS_DMC_01_set_position(gCardNo, node1, 0, 0);
            CPCI_DMC.CS_DMC_01_set_command(gCardNo, node1, 0, 0);
           // btnralm.Enabled = true;
            btnstop.Enabled = true;
           // btnreset1.Enabled = true;
            btnNmove.Enabled = true;
            btnPmove.Enabled = true;
           // chksvon.Checked = false;
           // chksvon.Enabled = true;

            excelTime = 0;
        }

        //private void chksvon_CheckedChanged(object sender, EventArgs e)
        //{
        //   // gIsServoOn = chksvon.Checked;
        //    gnodeid = ushort.Parse(cmbNodeID.Text);
        //    //btnWork.Enabled = true;
        //    rc = CPCI_DMC.CS_DMC_01_set_rm_04pi_ipulser_mode(gCardNo, node1, 0, 1);
        //    rc = CPCI_DMC.CS_DMC_01_set_rm_04pi_opulser_mode(gCardNo, node1, 0, 1);
        //    rc = CPCI_DMC.CS_DMC_01_ipo_set_svon(gCardNo, node1, 0, (ushort)(gIsServoOn ? 1 : 0));
        //}

        private void chart4_Click(object sender, EventArgs e)
        {

        }

        private void btnstop_Click(object sender, EventArgs e)
        {
            LogOutput("馬達緊急停止");
            rc = CPCI_DMC.CS_DMC_01_emg_stop(gCardNo, node1, 0);
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        public void OpenCard()
        {
            LogOutput("開啟軸卡中");
            
            ushort i, card_no = 0;
            btnstop.Enabled = false;
            btnNmove.Enabled = false;
            btnPmove.Enabled = false;

            for (i = 0; i < 4; i++)
            {
                SlaveTable[i] = 0;
            }
            txtSlaveNum.Text = "0";
            CmbCardNo.Items.Clear();
            cmbNodeID.Items.Clear();

            rc = CPCI_DMC.CS_DMC_01_open(ref existcard);

            if (existcard <= 0)
            {
                LogOutput("未發現軸卡");
                //MessageBox.Show("No DMC-NET card can be found!");
            }
            else
            {

                for (i = 0; i < existcard; i++)
                {
                    rc = CPCI_DMC.CS_DMC_01_get_CardNo_seq(i, ref card_no);
                    gCardNoList[i] = card_no;

                    CmbCardNo.Items.Insert(i, card_no);

                }
                CmbCardNo.SelectedIndex = 0;
                gCardNo = gCardNoList[0];

                for (i = 0; i < existcard; i++)
                {
                    rc = CPCI_DMC.CS_DMC_01_pci_initial(gCardNoList[i]);
                    if (rc != 0)
                    {
                        LogOutput("無法啟動軸卡");
                        MessageBox.Show("Can't boot PCI_DMC Master Card!");
                    }

                    rc = CPCI_DMC.CS_DMC_01_initial_bus(gCardNoList[i]);
                    if (rc != 0)
                    {
                        LogOutput("軸卡初始化失敗");
                        MessageBox.Show("Initial Failed!");
                    }
                    else
                    {
                        rc = CPCI_DMC.CS_DMC_01_start_ring(gCardNo, 0);                      //Start communication                      
                        rc = CPCI_DMC.CS_DMC_01_get_device_table(gCardNo, ref DeviceInfo);   //Get Slave Node ID 
                        rc = CPCI_DMC.CS_DMC_01_get_node_table(gCardNo, ref SlaveTable[0]);
                        LogOutput("成功與軸卡連線");
                    }
                }
            }
        }

        public void ServoRST()
        {
            LogOutput("伺服馬達歸零");
            gnodeid = ushort.Parse(cmbNodeID.Text);
            CPCI_DMC.CS_DMC_01_set_position(gCardNo, node1, 0, 0);
            CPCI_DMC.CS_DMC_01_set_command(gCardNo, node1, 0, 0);
            rc = CPCI_DMC.CS_DMC_01_set_rm_04pi_ipulser_mode(gCardNo, node1, 0, 1);
            rc = CPCI_DMC.CS_DMC_01_set_rm_04pi_opulser_mode(gCardNo, node1, 0, 1);
            btnstop.Enabled = true;
            btnNmove.Enabled = true;
            btnPmove.Enabled = true;
        }

        public void ServoON(bool IsOn)
        {
            switch (IsOn)
            {
                case true:
                    rc = CPCI_DMC.CS_DMC_01_ipo_set_svon(gCardNo, node1, 0, 1);
                    LogOutput("伺服馬達已啟動");
                    break;
                case false:
                    rc = CPCI_DMC.CS_DMC_01_ipo_set_svon(gCardNo, node1, 0, 0);
                    LogOutput("伺服馬達已關閉");
                    break;
            }
            ServoWorking = IsOn;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool check = false;
            try
            {
                OpenCard();
                FindSlave();
                ServoRST();
                btnWork.Enabled = true;
            }
            catch (Exception)
            {

            }
        }

        public void FindSlave(){
            LogOutput("開始搜尋控制器...");
            ushort i, lMask = 0x1, p = 0;
            uint DeviceType = 0, IdentityObject = 0;
            btnstop.Enabled = false;
            btnNmove.Enabled = false;
            btnPmove.Enabled = false;
            gNodeNum = 0;
            txtSlaveNum.Text = "0";
            cmbNodeID.Items.Clear();

            for (i = 0; i < 1; i++) NodeID[i] = 0;

            if (SlaveTable[0] == 0)
            {
                MessageBox.Show("CardNo: " + gCardNo.ToString() + " No slave found!");
                LogOutput("未發現控制器");
            }
            else
            {
                for (i = 0; i < 32; i++)
                {
                    if ((SlaveTable[0] & lMask) != 0)
                    {
                        NodeID[gNodeNum] = (ushort)(i + 1);
                        gNodeNum++;
                        rc = CPCI_DMC.CS_DMC_01_get_devicetype((short)gCardNo, (ushort)(i + 1), (ushort)0, ref DeviceType, ref IdentityObject);
                        if (rc != 0)
                        {
                            MessageBox.Show("get_devicetype failed - code=" + rc);
                        }
                        else
                        {
                            switch (DeviceType)
                            {
                                case 0x4020192:				//Servo A2 series
                                    cmbNodeID.Items.Add(i + 1);
                                    p++;
                                    break;
                                case 0x6020192:				//Servo M series
                                    cmbNodeID.Items.Add(i + 1);
                                    p++;
                                    break;
                                case 0x8020192:				//Servo A2R series
                                    cmbNodeID.Items.Add(i + 1);
                                    p++;
                                    break;
                                case 0x9020192:				//Servo S series
                                    cmbNodeID.Items.Add(i + 1);
                                    p++;
                                    break;
                            }
                        }
                    }
                    lMask <<= 1;
                }
                if (p == 0)
                {
                    MessageBox.Show("No A2 Servo Device Found!");
                }
                else
                {
                    txtSlaveNum.Text = gNodeNum.ToString();
                    cmbNodeID.SelectedIndex = 0;
                    LogOutput("控制卡連線完成");
                }
            }
        }

        

        private void working_PLC()
        {
            while (ServoWorking)
            {
                if (!ThWorking_PLC.IsAlive)
                {
                    //lblcount.Text = "exit";
                    break;
                }
                Send("000000000006" + "010313000004");
                Listen();
                MotorListen();
                chart1.Series[0].Points.AddXY(ktrRpm1.Count, ktrRpm1[ktrRpm1.Count - 1]);
                chart2.Series[0].Points.AddXY(ktrRpm2.Count, ktrRpm2[ktrRpm2.Count - 1]);
                chart3.Series[0].Points.AddXY(ktrTorque1.Count, ktrTorque1[ktrTorque1.Count - 1]);
                chart4.Series[0].Points.AddXY(ktrTorque2.Count, ktrTorque2[ktrTorque2.Count - 1]);
                chart5.Series[0].Points.AddXY(motorRpm1.Count, motorRpm1[motorRpm1.Count - 1]);
                chart6.Series[0].Points.AddXY(motorTorque1.Count, motorTorque1[motorTorque1.Count - 1]);
            }
        }

        private void Send(string Str)
        {
            byte[] A = new byte[1]; //初始需告陣列(因不知道資料大小，下面會做陣列調整)
            for (int i = 0; i < Str.Length / 2; i++)
            {
                Array.Resize(ref A, Str.Length / 2);  //Array.Resize(ref 陣列名稱, 新的陣列大小)  
                string str2 = Str.Substring(i * 2, 2);
                A[i] = Convert.ToByte(str2, 16); //字串依照"frombase"轉換數字(Byte)
            }
            T.Send(A, 0, Str.Length / 2, SocketFlags.None);
        }

        private void btnFinish_Click(object sender, EventArgs e)
        {
            btnFinish.Enabled = false;
            btnSaveExcel.Enabled = true;
            LogOutput("實驗結束");
            ServoON(false);
        }

        private void btnexit_Click(object sender, EventArgs e)
        {
            ServoCon.Enabled = false;
            btnWork.Enabled = false;
            btnFinish.Enabled = false;
            btnSaveExcel.Enabled = false;
            LogOutput("關閉伺服馬達....");
            rc = CPCI_DMC.CS_DMC_01_ipo_set_svon(gCardNo, node1, 0, 0);
            rc = CPCI_DMC.CS_DMC_01_reset_card(gCardNo);
            CPCI_DMC.CS_DMC_01_close();
            LogOutput("Bye bye~~~");
            Thread.Sleep(1500);
            Environment.Exit(Environment.ExitCode);

        }
        private void SetChartType()
        {
            SeriesChartType type = new SeriesChartType();

            switch (CType.ToString())
            {
                case "Line" : type = SeriesChartType.Line;
                    break;
                case "Spline": type = SeriesChartType.Spline;
                    break;

            }
            chart1.Series[0].ChartType = type;
            chart2.Series[0].ChartType = type;
            chart3.Series[0].ChartType = type;
            chart4.Series[0].ChartType = type;
            chart5.Series[0].ChartType = type;
            chart6.Series[0].ChartType = type;
        }

        private void CleanChart()
        {
            chart1.Series[0].Points.Clear();
            chart2.Series[0].Points.Clear();
            chart3.Series[0].Points.Clear();
            chart4.Series[0].Points.Clear();
            chart5.Series[0].Points.Clear();
            chart6.Series[0].Points.Clear();
            ktrRpm1.Clear();
            ktrRpm2.Clear();
            ktrTorque1.Clear();
            ktrTorque2.Clear();
            motorRpm1.Clear();
            motorTorque1.Clear();
        }

        private void btnWork_Click(object sender, EventArgs e)
        {
            btnWork.Enabled = false;
            btnFinish.Enabled = true;
            LogOutput("實驗開始");

            SetChartType();
            CleanChart();

            ServoON(true);

            double m_Tacc = Double.Parse(txtTacc.Text), m_Tdec = Double.Parse(txtTdec.Text);
            int m_Rpm = Int16.Parse(txtRpm2.Text)*10;
            gnodeid = ushort.Parse(cmbNodeID.Text);
            /* Set up Velocity mode parameter */
            rc = CPCI_DMC.CS_DMC_01_set_velocity_mode(gCardNo, node1, 0, m_Tacc, m_Tdec);
            //* Start Velocity move: rpm > 0 move forward , rpm < 0 move negative */
            rc = CPCI_DMC.CS_DMC_01_set_velocity(gCardNo, node1, 0, m_Rpm);


            
        }

        private void Listen()
        {
            EndPoint ServerEP = (EndPoint)T.RemoteEndPoint;
            byte[] B = new byte[1023];
            int inLen = 0;
            try
            {
                inLen = T.ReceiveFrom(B, ref ServerEP);
            }
            catch (Exception)
            {
                T.Close();
                MessageBox.Show("伺服器中斷連線!");
                btnConnectPLC.Enabled = true;
            }
            txtReceive.Text = BitConverter.ToString(B, 6, inLen - 6);
            string[] ary = txtReceive.Text.Split('-');
            //將讀取到的16進制碼換成10進制碼，且切割後的陣列兩個為1組
            double rpm1, rpm2, torque1, torque2;
            rpm1 = changeVoltage0x16(Int32.Parse(ary[3] + ary[4], System.Globalization.NumberStyles.HexNumber));
            rpm2 = changeVoltage0x16(Int32.Parse(ary[5] + ary[6], System.Globalization.NumberStyles.HexNumber));
            torque1 = changeVoltage0x16(Int32.Parse(ary[7] + ary[8], System.Globalization.NumberStyles.HexNumber));
            torque2 = changeVoltage0x16(Int32.Parse(ary[9] + ary[10], System.Globalization.NumberStyles.HexNumber));
            rpm1 = (rpm1*10/8000) * rpmRate1;
            rpm2 = (rpm2*10/8000) * rpmRate2;
            torque1 = (torque1 * 10 / 8000) * torqueRate_10;
            torque2 = (torque2 * 10 / 8000) * torqueRate_50;
            ktrRpm1.Add(rpm1);
            ktrRpm2.Add(rpm2);
            ktrTorque1.Add(torque1);
            ktrTorque2.Add(torque2);

        }

        private void MotorListen()
        {
            rc = CPCI_DMC.CS_DMC_01_get_rpm(gCardNo, node1, 0, ref spd1);
            if (rc == 0)
            {
                txtspeed1.Text = spd1.ToString();
            }
            //Torque
            rc = CPCI_DMC.CS_DMC_01_get_torque(gCardNo, node1, 0, ref toe1);
            if (rc == 0)
            {
                //扭矩是千分比
                txtTorque1.Text = ((double)toe1 / 1000 * 7.16).ToString();
            }
            motorTorque1.Add((double)toe1 / 1000 * 7.16);
            motorRpm1.Add(spd1 / 10);
        }
        public double changeVoltage0x16(double v)
        {
            if (v > 32767)
                return ((65535 - v + 1) * (-1));
            else
                return v;
        }
    }
}
