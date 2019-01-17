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

using PCI_DMC;
using PCI_DMC_ERR;


namespace KTR_Measure
{
    
    public partial class Form1 : Form
    {
        
        public const int rpmRate1 = 400; //ktr比例
        public const int rpmRate2 = 4;
        public const int torqueRate = 5;
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
        bool gIsServoOn;
        short spd1 = 0, spd2 = 0, toe1 = 0, toe2 = 0;
        bool ServoWorking = false;
        TextBox[] txtIoSts = new TextBox[16];


        List<double> ktrTorque1 = new List<double>();
        List<double> ktrTorque2 = new List<double>();
        List<double> ktrRpm1 = new List<double>();
        List<double> ktrRpm2 = new List<double>();
        List<double> motorTorque1 = new List<double>();
        //List<double> motorTorque2 = new List<double>();
        List<double> motorRpm1 = new List<double>();
        //List<double> motorRpm2 = new List<double>();
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


        private void btnConnectPLC_Click(object sender, EventArgs e)
        {
            dialog.AppendText("開始連線至PLC\n");
            string IP = txtIPToPLC.Text;                //設定變數IP，其字串
            int Port = int.Parse(txtPortToPLC.Text);    //設定變數Port，為整數
            try
            {
                //IPAddress是IP，如" 127.0.0.1"  ;IPEndPoint是ip和端口對的組合，如"127.0.0.1: 1000 "  
                IPEndPoint EP = new IPEndPoint(IPAddress.Parse(IP), Port);
                //new Socket( 通訊協定家族IP4 , 通訊端類型 , 通訊協定TCP)
                T = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                T.Connect(EP); //建立連線
                lblConnectStatus.Text = "已連線至PLC";
                dialog.AppendText("已連線至PLC\n");
                ServoCon.Enabled = true;
            }
            catch (Exception)
            {
                lblConnectStatus.Text = "無法連線至PLC,請檢查線路或IP";
                dialog.AppendText("無法連線至PLC,請檢查線路或IP\n");
                return;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;
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
            dialog.AppendText("檔案儲存中.....\n");
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
            dialog.AppendText("檔案已儲存至：\n");
            dialog.AppendText(FileStr);
            dialog.AppendText(".xlsx\n");
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

        private void chksvon_CheckedChanged(object sender, EventArgs e)
        {
           // gIsServoOn = chksvon.Checked;
            gnodeid = ushort.Parse(cmbNodeID.Text);
            //btnWork.Enabled = true;
            rc = CPCI_DMC.CS_DMC_01_set_rm_04pi_ipulser_mode(gCardNo, node1, 0, 1);
            rc = CPCI_DMC.CS_DMC_01_set_rm_04pi_opulser_mode(gCardNo, node1, 0, 1);
            rc = CPCI_DMC.CS_DMC_01_ipo_set_svon(gCardNo, node1, 0, (ushort)(gIsServoOn ? 1 : 0));
        }

        private void chart4_Click(object sender, EventArgs e)
        {

        }

        private void btnstop_Click(object sender, EventArgs e)
        {
            dialog.AppendText("馬達緊急停止\n");
            rc = CPCI_DMC.CS_DMC_01_emg_stop(gCardNo, node1, 0);
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        public void OpenCard()
        {
            dialog.AppendText("開啟軸卡中\n");
            
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
                dialog.AppendText("未發現軸卡\n");
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
                        dialog.AppendText("無法啟動軸卡\n");
                        MessageBox.Show("Can't boot PCI_DMC Master Card!");
                    }

                    rc = CPCI_DMC.CS_DMC_01_initial_bus(gCardNoList[i]);
                    if (rc != 0)
                    {
                        dialog.AppendText("軸卡初始化失敗\n");
                        MessageBox.Show("Initial Failed!");
                    }
                    else
                    {
                        rc = CPCI_DMC.CS_DMC_01_start_ring(gCardNo, 0);                      //Start communication                      
                        rc = CPCI_DMC.CS_DMC_01_get_device_table(gCardNo, ref DeviceInfo);   //Get Slave Node ID 
                        rc = CPCI_DMC.CS_DMC_01_get_node_table(gCardNo, ref SlaveTable[0]);
                        dialog.AppendText("成功與軸卡連線\n");
                    }
                }
            }
        }

        public void ServoRST()
        {
            dialog.AppendText("伺服馬達歸零\n");
            gnodeid = ushort.Parse(cmbNodeID.Text);
            CPCI_DMC.CS_DMC_01_set_position(gCardNo, node1, 0, 0);
            CPCI_DMC.CS_DMC_01_set_command(gCardNo, node1, 0, 0);
            btnstop.Enabled = true;
            btnNmove.Enabled = true;
            btnPmove.Enabled = true;
        }

        public void ServoON()
        {
            dialog.AppendText("伺服馬達以啟動\n");
            gnodeid = ushort.Parse(cmbNodeID.Text);
            rc = CPCI_DMC.CS_DMC_01_set_rm_04pi_ipulser_mode(gCardNo, node1, 0, 1);
            rc = CPCI_DMC.CS_DMC_01_set_rm_04pi_opulser_mode(gCardNo, node1, 0, 1);
            rc = CPCI_DMC.CS_DMC_01_ipo_set_svon(gCardNo, node1, 0, 1);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool check = false;
            try
            {
                OpenCard();
                FindSlave();
                ServoRST();
                ServoON();
                btnWork.Enabled = true;
            }
            catch (Exception)
            {

            }
            
            

        }

        public void FindSlave(){
            dialog.AppendText("開始搜尋控制器..\n");
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
                dialog.AppendText("未發現控制器\n");
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
                    dialog.AppendText("控制卡連線完成\n");
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
            dialog.AppendText("實驗結束 馬達停止\n");
            rc = CPCI_DMC.CS_DMC_01_set_velocity_stop(gCardNo, node1, 0, 1);
            ServoWorking = false;
        }

        private void btnexit_Click(object sender, EventArgs e)
        {
            ServoCon.Enabled = false;
            btnWork.Enabled = false;
            btnFinish.Enabled = false;
            btnSaveExcel.Enabled = false;
            dialog.AppendText("關閉伺服馬達....\n");
            rc = CPCI_DMC.CS_DMC_01_ipo_set_svon(gCardNo, node1, 0, 0);
            rc = CPCI_DMC.CS_DMC_01_reset_card(gCardNo);
            CPCI_DMC.CS_DMC_01_close();
            dialog.AppendText("Bye bye~~~\n");
            Thread.Sleep(1500);
            Environment.Exit(Environment.ExitCode);

        }

        private void btnWork_Click(object sender, EventArgs e)
        {
            btnWork.Enabled = false;
            btnFinish.Enabled = true;
            dialog.AppendText("實驗開始\n");
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

            ServoWorking = true;

            double m_Tacc = Double.Parse(txtTacc.Text), m_Tdec = Double.Parse(txtTdec.Text);
            int m_Rpm = Int16.Parse(txtRpm2.Text)*10;
            gnodeid = ushort.Parse(cmbNodeID.Text);
            /* Set up Velocity mode parameter */
            rc = CPCI_DMC.CS_DMC_01_set_velocity_mode(gCardNo, node1, 0, m_Tacc, m_Tdec);
            //* Start Velocity move: rpm > 0 move forward , rpm < 0 move negative */
            rc = CPCI_DMC.CS_DMC_01_set_velocity(gCardNo, node1, 0, m_Rpm);


            ThWorking_PLC = new Thread(working_PLC);
            ThWorking_PLC.Start();
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
            torque1 = (torque1 * 10 / 8000) * torqueRate;
            torque2 = (torque2 * 10 / 8000) * torqueRate;
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
