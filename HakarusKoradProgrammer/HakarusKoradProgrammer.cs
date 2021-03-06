﻿using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using System.IO.Ports;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;



namespace HakarusKoradProgrammer
{
    public partial class HakarusKoradProgrammer : Form
    {
        List<SerialDevice> _DeviceList = new List<SerialDevice>();    //List to store the Devices
        List<TestSequenceElement> _TestSequenceElements = new List<TestSequenceElement>(); //List to store the test sequences
        List<TestSequenceElement> _DataLoggingList = new List<TestSequenceElement>(); //List to store the samples logged during a test run

        XmlSerializer _XmlSerial;//creates ana XmlSerializer object for use in the XML IO sections


        private int[] _AcceptableBaudRates = { 110, 300, 600, 1200, 2400, 4800, 9600, 14400, 19200, 38400, 57600, 115200, 128000};

        private float _voltage = 0;
        private float _current = 0;
        private float _power = 0;
        private float _resistance = 0;

        private bool _ThreadEnd = false;
        private bool _Beeper = false;
        private bool _IsTesting = false;

        private string FilePath = "";
        private string FileName = "";

        //Sets teh column width of the dataLoging and componentTest list boxes
        private string lvlvComponentTesterBox = "{0,0}{1,35}{2,70}";
        private string LoggingLayout = "{0,0}{1,25}{2,25}{3,35}{4,35}";

        Thread PowerMeasurement;


        #region initialization
        public HakarusKoradProgrammer()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            #region Text box initialisation
            //Populates the form text boxes
            cbbParity.Items.Add("None");
            cbbParity.Items.Add("Odd");
            cbbParity.Items.Add("Even");
            cbbParity.Items.Add("Mark");
            cbbParity.Items.Add("Space");

            cbbStopBits.Items.Add("One");
            cbbStopBits.Items.Add("OnePointFive");
            cbbStopBits.Items.Add("Two");

            cbbFlowControl.Items.Add("None");
            cbbFlowControl.Items.Add("RequestToSend");
            cbbFlowControl.Items.Add("RequestToSendXOnXOff");
            cbbFlowControl.Items.Add("XOnXOff");

            txtConnected.Text = "No";


            txtPowerState.Text = "Powered Off";

            txtListVoltage.Text = "0";
            txtListCurrent.Text = "0";
            txtListTime.Text = "0";

            txtVoltage.Text = "0";
            txtCurrent.Text = "0";
            #endregion
            lbxTestSequence.MultiColumn = true;
            lbxTestSequence.Items.Add(string.Format(lvlvComponentTesterBox, "Voltage", "Current", "Time"));
            ComPortInit();
        }
        private void PollingThreadGen()//Generates the thread for polling the power related values from the current selected device
        {
            Thread.Sleep(400);
            //This spawns a thread that will poll the currently selected device and set the actual volt/current output txtbox
            PowerMeasurement = new Thread(PowerPoller);//The new thread (PowerMeasurement) will run the PowerPoller method
            Console.WriteLine("Thread created");



            PowerMeasurement.IsBackground = true;//Sets the thread to a background thread that can run off and do its own thing independantly


            PowerMeasurement.Start();//actually starts the threads execution

            Console.WriteLine("Threads started");
        }
        #endregion


        #region Connection Settings
        private void btnCreateDevice_Click(object sender, EventArgs e)
        {
            CreateDevice();
        }
        private void CreateDevice()
        {
            if (string.IsNullOrWhiteSpace(txtDeviceName.Text))
            {
                Console.WriteLine("Cannot have an empty Device name");
                MessageBox.Show("Cannot have an empty Device name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            else if (string.IsNullOrWhiteSpace(cbbComPort.Text))
            {
                Console.WriteLine("Cannot have an empty COM port");
                MessageBox.Show("Cannot have an empty COM port", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                DeviceListValidation();
            }
        }

        private void btnRemoveDevice_Click(object sender, EventArgs e)
        {

            RemoveDevice();
        }
        private void RemoveDevice()
        {
            int index = 0;
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == cbbDevice.Text)
                {
                    if (!device._isConnected)
                    {
                        device.Disconnect();
                        _ThreadEnd = true;
                        Console.WriteLine("Removing device {0}", device._deviceName);
                        cbbDevice.SelectedIndex = -1;
                        cbbDevice.Items.Remove(device._deviceName.ToString());
                        _DeviceList.RemoveAt(index);
                        txtDeviceID.Text = "";
                        txtConnected.Text = "False";
                        return;
                    }
                    else
                    {
                        Console.WriteLine("Removing device {0}", device._deviceName);
                        _ThreadEnd = true;
                        cbbDevice.SelectedIndex = -1;
                        cbbDevice.Items.Remove(device._deviceName.ToString());
                        _DeviceList.RemoveAt(index);
                        txtDeviceID.Text = "";
                        txtConnected.Text = "False";

                        return;
                    }

                }
                index += 1;
            }
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            CreateDevice();
            Connect();
        }
        private void Connect()
        {
            _ThreadEnd = false;
            if (IsDeviceSelected())
            {
                Console.WriteLine("Connecting device...");
                foreach (SerialDevice device in _DeviceList)
                {
                    if (device._deviceName == cbbDevice.Text)
                    {
                        if (device._isConnected)
                        {
                            //Pop up device already connected message
                            DeviceConnected();
                        }
                        else
                        {

                            Console.WriteLine("Connecting {0} to port {1}", cbbDevice.Text, cbbComPort.Text);
                            txtConnected.Text = device.Connect().ToString();
                            //If statement safeguards the code so if there is a failed connection (even if the device was detected) the program wont 
                            //go on to do the Initialisation of the device settings which requires port access which would cause a program crash
                            if (txtConnected.Text == "True")
                            {
                                GetDeviceId();
                                PollingThreadGen();
                            }
                            else
                            {
                                return;
                            }


                            //Initialises the device settings
                            device.SendQueuePush("OUT", "0");
                            device.SendQueuePush("VSET1:", "0");
                            device.SendQueuePush("ISET1:", "0");
                            device.SendQueuePush("BEEP", "1");
                            Thread.Sleep(100);
                        }
                    }

                }
            }
            else
            {
                DeviceNotSelected();
            }
        }

        private void btnDisconnect_Click(object sender, EventArgs e)
        {
            disablePower();
            Thread.Sleep(100);
            Disconnect();
            RemoveDevice();
        }
        private void Disconnect()
        {
            if (IsDeviceSelected())
            {
                foreach (SerialDevice device in _DeviceList)
                {
                    if (device._deviceName == cbbDevice.Text)
                    {
                        if (device._isConnected)
                        {
                            _ThreadEnd = true;
                            //Thread.Sleep(100);
                            Console.WriteLine("Closing connection to {0}.", device._comPort);

                            device.Disconnect();
                            txtConnected.Text = device.Disconnect().ToString();
                            PowerMeasurement.Abort();
                            if (PowerMeasurement.IsAlive)
                            {
                                Console.WriteLine("Thread not terminated");
                            }
                            else
                            {
                                Console.WriteLine("Thread terminated");
                            }
                        }
                        else
                        {
                            DeviceNotConnected();
                        }
                    }
                }
            }
            else
            {
                DeviceNotSelected();
            }
            
            Console.WriteLine("Device fully disconnected");
        }

        private void DeviceListValidation()
        {
            Console.WriteLine("Device name entered:");
            if (_DeviceList.Any())
            {
                if (_AcceptableBaudRates.Contains(Int32.Parse(txtBaudRate.Text)))
                {
                    Console.Write("List not Empty, checking to see if device is in the list:...");
                    foreach (SerialDevice device in _DeviceList)
                    {
                        if (device._deviceName == txtDeviceName.Text)
                        {
                            Console.WriteLine("Device already exists");
                            MessageBox.Show("Device already exits", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Console.WriteLine();
                        }
                        else if (device._comPort == cbbComPort.Text)
                        {
                            Console.WriteLine("COM port already in use");
                            MessageBox.Show("COM port already in use", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Console.WriteLine();
                        }
                        else
                        {
                            Console.WriteLine("Device List is not Empty, adding device");

                            AddDevice(cbbComPort.Text, Int32.Parse(txtBaudRate.Text), Int32.Parse(txtDataBits.Text), cbbParity.Text, cbbStopBits.Text, cbbFlowControl.Text, txtDeviceName.Text);
                            _ThreadEnd = false;
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Please enter an acceptable baud rate");
                    MessageBox.Show("Please enter an acceptable baud rate", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                if(_AcceptableBaudRates.Contains(Int32.Parse(txtBaudRate.Text)))
                {
                    Console.WriteLine("Device List is Empty, adding device");
                    Console.WriteLine();
                    AddDevice(cbbComPort.Text, Int32.Parse(txtBaudRate.Text), Int32.Parse(txtDataBits.Text), cbbParity.Text, cbbStopBits.Text, cbbFlowControl.Text, txtDeviceName.Text);
                }
                else
                {
                    Console.WriteLine("Please enter an acceptable baud rate (110, 300, 600, 1200, 2400, 4800, 9600, 14400, 19200, 38400, 57600, 115200, 128000, 256000)");
                    MessageBox.Show("Please enter an acceptable baud rate \n (110, 300, 600, 1200, 2400, 4800, 9600, 14400, 19200, 38400, 57600, 115200, 128000, 256000)", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void AddDevice(String comPort, int baudRat, int dataBits, string parity, string stopBits, string flowControl, string deviceName)
        {
            if(dataBits <6 | dataBits  > 8)
            {
                Console.WriteLine("Data bits are out of range, value must be between 6 and 8");
                Console.WriteLine();
            }
            else
            {
                SerialDevice device = new SerialDevice(comPort, baudRat, dataBits, parity, stopBits, flowControl, deviceName);
                _DeviceList.Add(device);
                cbbDevice.Items.Add(deviceName);
                cbbDevice.SelectedIndex = cbbDevice.FindString(deviceName);
            }
        }
        //Sets the txtDeviceId txtbox to the current device (happens when device is connected)
        private void GetDeviceId()
        {
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == cbbDevice.Text)
                {
                    Thread.Sleep(150);
                    device.SendQueuePush("*IDN?", "");
                    Thread.Sleep(125);
                    txtDeviceID.Text = device._IDN;
                    Console.WriteLine("Device connected, ID: {0}", txtDeviceID.Text);
                }
            }

        }
        #endregion

        #region Voltage / Current control
        private void txtApplyVoltage_Click(object sender, EventArgs e)
        {
            bool _isDeviceSelected = IsDeviceSelected();
            if(_isDeviceSelected)
            {
                if (string.IsNullOrWhiteSpace(txtVoltage.Text))
                {
                    Console.WriteLine("Please enter a voltage");
                    MessageBox.Show("Error, Please enter a voltage", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    if (float.Parse(txtVoltage.Text) > 31 | float.Parse(txtVoltage.Text) < 0)
                    {
                        Console.WriteLine("Entered voltage is out of range, Voltage has to be between 0 and 31");
                        MessageBox.Show("Error, please enter a valid voltage (0-31'V)", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        Console.WriteLine("Applying voltage {0}", txtVoltage.Text);
                        foreach (SerialDevice device in _DeviceList)
                        {
                            if (device._deviceName == cbbDevice.Text)
                            {

                                device.SendQueuePush("VSET1:", txtVoltage.Text);
                                txtSetVoltage.Clear();
                                txtSetVoltage.Text = String.Format("{0:00.00}", txtVoltage.Text);
                            }
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("Please select a device");
                DeviceNotSelected();
            }

        }
        private void txtApplyCurrent_Click(object sender, EventArgs e)
        {
            bool _isDeviceSelected = IsDeviceSelected();
            if (_isDeviceSelected)
            {
                if (string.IsNullOrWhiteSpace(txtCurrent.Text))
                {
                    Console.WriteLine("Please enter a Current");
                    MessageBox.Show("Error, Please enter a Current", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    if (float.Parse(txtCurrent.Text) > 5 | float.Parse(txtCurrent.Text) < 0)
                    {
                        Console.WriteLine("Entered Current is out of range, Current has to be between 0 and 3");
                        MessageBox.Show("Error, please enter a valid Current (0-5A')", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        Console.WriteLine("Applying Current {0}", txtCurrent.Text);
                        foreach (SerialDevice device in _DeviceList)
                        {
                            if (device._deviceName == cbbDevice.Text)
                            {
                                device.SendQueuePush("ISET1:", txtCurrent.Text);
                                txtSetCurrent.Clear();
                                txtSetCurrent.Text = String.Format("{0:0.000}", txtCurrent.Text);
                            }
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("Device not selected");
                DeviceNotSelected();
            }
        }
        private void btnApplyBoth_Click(object sender, EventArgs e)
        {
            bool _isDeviceSelected = IsDeviceSelected();
            if (_isDeviceSelected)
            {
                if (string.IsNullOrWhiteSpace(txtCurrent.Text))
                {
                    Console.WriteLine("Please enter a Current");
                    MessageBox.Show("Error, Please enter a Current", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (string.IsNullOrWhiteSpace(txtVoltage.Text))
                {
                    Console.WriteLine("Please enter a voltage");
                    MessageBox.Show("Error, Please enter a voltage", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    if (float.Parse(txtCurrent.Text) > 5 | float.Parse(txtCurrent.Text) < 0 && float.Parse(txtVoltage.Text) > 31 | float.Parse(txtVoltage.Text) < 0)
                    {
                        Console.WriteLine("Entered Current/Voltage is out of range");
                        MessageBox.Show("Error, please enter a valid Current/Voltage", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        Console.WriteLine("Applying Current {0}A and voltage {1}V", txtCurrent.Text, txtVoltage.Text);
                        foreach (SerialDevice device in _DeviceList)
                        {
                            if (device._deviceName == cbbDevice.Text)
                            {
                                device.SendQueuePush("ISET1:", txtCurrent.Text);
                                txtSetCurrent.Clear();
                                txtSetCurrent.Text = String.Format("{0:0.000}", txtCurrent.Text);
                                Thread.Sleep(10);
                                device.SendQueuePush("VSET1:", txtVoltage.Text);
                                txtSetVoltage.Clear();
                                txtSetVoltage.Text = String.Format("{0:00.00}", txtVoltage.Text);
                            }
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("Device not selected");
                DeviceNotSelected();
            }
        }
        private void btnPower_Click(object sender, EventArgs e)
        {
            //_PollThreadhold = true;
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == cbbDevice.Text)
                {
                    if (device._isConnected)
                    {
                        Thread.Sleep(200);
                        Console.WriteLine("Switching {0}'s power to {1}", device._deviceName, device.PowerSwitch());
                        if (device._isOn)
                        {
                            txtPowerState.Text = "Powered On";
                        }
                        else
                        {
                            txtPowerState.Text = "Powered Off";
                        }

                    }
                    else
                    {
                        DeviceNotConnected();
                    }

                }
            }
            //_PollThreadhold = false;
        }

        private void disablePower()
        {
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == cbbDevice.Text)
                {
                    if (device._isConnected)
                    {
                        if (device._isOn)
                        {
                            txtPowerState.Text = "Powered Off";
                            device.PowerSwitch();
                        }

                    }
                    else
                    {
                        DeviceNotConnected();
                    }

                }
            }

        }
        #endregion

        #region MessageBox errors
        public static void DeviceNotConnected()
        {
            MessageBox.Show("Error, device not connected", "Error: No connection", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Console.WriteLine("Error, device not connected");
        }
        public static void DeviceConnected()
        {
            MessageBox.Show("Error, device is already connected", "Error: Already connection", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Console.WriteLine("Error, device is already connected");
        }
        public static void DeviceNotSelected()
        {
            MessageBox.Show("Error, device is Not Selected", "Error: No device selected", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Console.WriteLine("Error, device is Not Selected");
        }



        #endregion

        
        #region VoltageCurrent update
        private void button1_Click(object sender, EventArgs e)
        {
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == cbbDevice.Text)
                {
                    device.SendQueuePush("VOUT1?", "");
                }
            }
        }

        private void btnMeasureCurrent_Click(object sender, EventArgs e)
        {
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == cbbDevice.Text)
                {
                    device.SendQueuePush("IOUT1?", "");
                }
            }
        }
        #endregion



        #region Form closing
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            disablePower();
            Thread.Sleep(50);
            _ThreadEnd = true;
            
        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            disablePower();
            Thread.Sleep(50);
            _ThreadEnd = true;
            
        }

        #endregion

        #region (OC/OV)Protection
        private void btnOvpOn_Click(object sender, EventArgs e)
        {
            foreach(SerialDevice device in _DeviceList)
            {
                if(device._deviceName == txtDeviceName.Text)
                {
                    device.SendQueuePush("OVP", "1");
                }
            }
        }

        private void btnOvpOff_Click(object sender, EventArgs e)
        {
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == txtDeviceName.Text)
                {
                    device.SendQueuePush("OVP", "0");
                }
            }
        }

        private void btnOcpOn_Click(object sender, EventArgs e)
        {
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == txtDeviceName.Text)
                {
                    device.SendQueuePush("OCP", "1");
                }
            }
        }

        private void btnOcpOff_Click(object sender, EventArgs e)
        {
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == txtDeviceName.Text)
                {
                    device.SendQueuePush("OCP", "0");
                }
            }
        }
        #endregion
        #region Load/Store
        private void txtLoadM1_Click(object sender, EventArgs e)
        {
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == txtDeviceName.Text)
                {
                    device.SendQueuePush("RCL", "1");
                }
            }
        }
        private void txtLoadM2_Click(object sender, EventArgs e)
        {
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == txtDeviceName.Text)
                {
                    device.SendQueuePush("RCL", "2");
                }
            }
        }
        private void txtLoadM3_Click(object sender, EventArgs e)
        {
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == txtDeviceName.Text)
                {
                    device.SendQueuePush("RCL", "3");
                }
            }
        }
        private void txtLoadM4_Click(object sender, EventArgs e)
        {
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == txtDeviceName.Text)
                {
                    device.SendQueuePush("RCL", "4");
                }
            }
        }
        private void txtLoadM5_Click(object sender, EventArgs e)
        {
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == txtDeviceName.Text)
                {
                    device.SendQueuePush("RCL", "5");
                }
            }
        }

        private void txtSaveM1_Click(object sender, EventArgs e)
        {
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == txtDeviceName.Text)
                {
                    device.SendQueuePush("SAV", "1");
                }
            }
        }
        private void txtSaveM2_Click(object sender, EventArgs e)
        {
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == txtDeviceName.Text)
                {
                    device.SendQueuePush("SAV", "2");
                }
            }
        }
        private void txtSaveM3_Click(object sender, EventArgs e)
        {
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == txtDeviceName.Text)
                {
                    device.SendQueuePush("SAV", "3");
                }
            }
        }
        private void txtSaveM4_Click(object sender, EventArgs e)
        {
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == txtDeviceName.Text)
                {
                    device.SendQueuePush("SAV", "4");
                }
            }
        }
        private void txtSaveM5_Click(object sender, EventArgs e)
        {
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == txtDeviceName.Text)
                {
                    device.SendQueuePush("SAV", "5");
                }
            }
        }

        #endregion
        #region Console
        private void btnSendMsg_Click(object sender, EventArgs e)
        {
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == txtDeviceName.Text)
                {
                    string result;
                    Console.WriteLine("Piping {0} to the device", txtCslIn.Text);
                    device.SendQueuePush("CSL", txtCslIn.Text);
                    Thread.Sleep(250);
                    result = device.ReturnConsole();
                    Console.WriteLine(device.ReturnConsole());
                    lbConsoleOut.Items.Add("Received back: ");
                    lbConsoleOut.Items.Add(result);
                    
                    //lbxTestSequence.Items.Add(string.Format(lvlvComponentTesterBox, "Voltage: " + voltage + "V", "Current: " + current + "A", "Time: " + time + "ms"));
                }
            }
            
        }
        #endregion
        #region Test sequencer
        private void btnAddList_Click(object sender, EventArgs e)
        {
            AddTestElement();
        }

        private void btnAddRamp_Click(object sender, EventArgs e)
        {
            AddRamp();
        }
        
        private void AddRamp()
        {
            if (!string.IsNullOrWhiteSpace(txtRampVFloor.Text) | !string.IsNullOrWhiteSpace(txtRampVCeiling.Text) | !string.IsNullOrWhiteSpace(txtRampCurrent.Text) | !string.IsNullOrWhiteSpace(txtRampTime.Text))
            {
                //lbxTestSequence.Items.Add(string.Format("Voltage: ", "Current: ", "Time: "));
                string svoltageFloor = txtRampVFloor.Text;
                string svoltageLim = txtRampVCeiling.Text;
                string scurrent = txtRampCurrent.Text;
                string stime = txtRampTime.Text;

                float fvfloor = float.Parse(svoltageFloor);
                float fvlim = float.Parse(svoltageLim);
                float fcurrent = float.Parse(scurrent);
                float ftime = float.Parse(stime);

                ftime = ftime * 1000;//converts the inputs time to ms
                stime = ftime.ToString();

                //number of ms between points used for the ramp
                const int timeInterval = 200;//Runs at 4hz (4 operation over a period of 1 second)

                int numberOfIntervals = (int)Math.Round((ftime / timeInterval), 0);
                Console.WriteLine("Number of intervals {0}", numberOfIntervals);

                float intervalDeltaVoltage = (fvlim - fvfloor) / numberOfIntervals;

                float currentVoltage = 0;


                for (int i = 0; i < numberOfIntervals + 1; i++)
                {
                    Console.WriteLine("calculated voltage is {0}", currentVoltage = fvfloor + (i * intervalDeltaVoltage));//calculates the current voltage 
                    lbxTestSequence.Items.Add(string.Format(lvlvComponentTesterBox, Math.Round(currentVoltage,2).ToString() + "V", scurrent + "A",(timeInterval * i) + "ms"));
                    Console.WriteLine("Added element to list box");
                    TestSequenceElement TestElement = new TestSequenceElement(Math.Round(currentVoltage, 2).ToString(), scurrent, (timeInterval).ToString());
                    _TestSequenceElements.Add(TestElement);
                }
                if(cbDeRamp.Checked)
                {
                    Console.WriteLine("Adding a deramp buffer object");
                    lbxTestSequence.Items.Add(string.Format(lvlvComponentTesterBox, 0.ToString() + "V", scurrent + "A", timeInterval * 2 + "ms"));
                    TestSequenceElement TestElement = new TestSequenceElement(0.ToString(), scurrent, (timeInterval * 2).ToString());
                    _TestSequenceElements.Add(TestElement);
                }                
            }
            else
            {
                MessageBox.Show("Please enter valid: Voltage, Current and time peramiters", "Error");
            }
        }
        private void AddTestElement()
        {
            //If the voltage, current and time input boxes are not empty
            if (!string.IsNullOrWhiteSpace(txtListVoltage.Text) | !string.IsNullOrWhiteSpace(txtListCurrent.Text) | !string.IsNullOrWhiteSpace(txtListTime.Text))
            {
                string voltage = txtListVoltage.Text;
                string current = txtListCurrent.Text;
                string time = txtListTime.Text;
                float ftime = float.Parse(time);
                ftime = ftime * 1000;
                time = ftime.ToString();

                lbxTestSequence.Items.Add(string.Format(lvlvComponentTesterBox, "Voltage: " + voltage + "V", "Current: " + current + "A", "Time: " + time + "ms"));

                TestSequenceElement TestElement = new TestSequenceElement(voltage, current, time);
                _TestSequenceElements.Add(TestElement);
            }
            else
            {
                MessageBox.Show("Please enter valid: Voltage, Current and time peramiters", "Error");
            }
        }

        private void btnRemoveList_Click(object sender, EventArgs e)
        {
            //2 is the number of items in the list box where its considered empty
            if(lbxTestSequence.Items.Count >= 2)
            {
                lbxTestSequence.Items.RemoveAt(lbxTestSequence.Items.Count - 1);
                lbxTestSequence.Refresh();
                _TestSequenceElements.RemoveAt(lbxTestSequence.Items.Count - 1);
            }
        }
        private void btnTest_Click(object sender, EventArgs e)  
        {
            StartTest();
        }
        private void StartTest()
        {
            lbLoggedData.Items.Clear();

            Thread TestThread = new Thread(TestExec);
            
            Console.WriteLine("Test thread created");

            TestThread.IsBackground = true;

            _IsTesting = true;

            
            TestThread.Start();
            Console.WriteLine("Test thread started");
        }

        private void TestExec()
        {
            //Initialises the device for testing
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == txtDeviceName.Text)
                {
                    device.SendQueuePush("VSET1:", "0");
                    device.SendQueuePush("ISET1:", "0");
                    device.SendQueuePush("OUT", "1");
                }
            }
            Thread.Sleep(750);//This allows the device to activate and zero out before the test begins
            Thread LoggingThread = new Thread(GraphPolling);
            LoggingThread.IsBackground = true;
            LoggingThread.Start();
            //Runs the tests
            for (int index = 0; index < _TestSequenceElements.Count; index++)
            {
                foreach (SerialDevice device in _DeviceList)
                {
                    if (device._deviceName == txtDeviceName.Text)
                    {
                        device.SendQueuePush("VSET1:", _TestSequenceElements[index].GetVoltage().ToString());
                        device.SendQueuePush("ISET1:", _TestSequenceElements[index].GetCurrent().ToString());
                        Thread.Sleep(  _TestSequenceElements[index].GetTime());
                    }
                }
            }
            //Deinitialises the device for testing, resetting back to stable values
            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == txtDeviceName.Text)
                {
                    device.SendQueuePush("VSET1:", "0");
                    device.SendQueuePush("ISET1:", "0");
                    device.SendQueuePush("OUT", "0");
                }
            }
            Thread.Sleep(100);
            //Ends the test thread
            _IsTesting = false;

            Console.WriteLine("Test thread ending");
        }
        #endregion

        private bool IsDeviceSelected()
        {
            //Just a method that returns true if there is a device selected or false if not
            if (string.IsNullOrWhiteSpace(cbbDevice.Text))
            {
                Console.WriteLine("No device selected");
                return false;
            }
            else
            {
                return true;
            }

        }
        #region Power monitoring
        private void PowerPoller()
        {
            Thread.Sleep(300);
            while(!_ThreadEnd)
            {
                PowerUpdater();
                Thread.Sleep(50);
            }
            if(_ThreadEnd)
            {
                Console.WriteLine("Thread going of of scope");
            }
        }
        private void PowerUpdater()
        {
            //if(string.IsNullOrWhiteSpace(cbbDevice.Text))
            //{
                foreach(SerialDevice device in _DeviceList)
                {
                    if(device._deviceName == txtDeviceName.Text)
                    {
                    txtActiveVoltage.Text = device._fvoltage.ToString();
                    txtActiveCurrent.Text = device._fcurrent.ToString();
                    if(device._current == null | device._voltage == null)
                        {

                        }
                        else
                        {
                            
                            _voltage = device._fvoltage;
                            _current = device._fcurrent;

                            _power = _voltage * _current;
                            _resistance = _voltage / _current;
                            Math.Round(_resistance);

                            txtActivePower.Text = _power.ToString();
                            txtResistance.Text = _resistance.ToString(); 
                            
                        }
                    }
                }
           // }
        }

        #endregion

        #region GraphPolling
        private void GraphPolling()
        {
            int index = 0;
            Console.WriteLine("IS testing: {0}",_IsTesting);

            //System timer to get timestamps for the events
            Stopwatch stopWatch = new Stopwatch();

            //Initialises a clean list for logging data into
            List<TestSequenceElement> DataLoggingList = new List<TestSequenceElement>();

            //Console.WriteLine("Entering the Data Logger");

            stopWatch.Start();
            while (_IsTesting)
            {
                if(_IsTesting)
                {
                    index = index + 1;
                    //Console.WriteLine("Testing");
                    //Console.WriteLine("Adding {0}V, {1}A, {2}W, {3}Ohms", txtActiveVoltage.Text, txtActiveCurrent.Text, txtActivePower.Text, txtResistance.Text);
                    TestSequenceElement TestSnippet = new TestSequenceElement(txtActiveVoltage.Text, txtActiveCurrent.Text, txtActivePower.Text, txtResistance.Text, stopWatch.ElapsedMilliseconds);
                    DataLoggingList.Add(TestSnippet);
                    Thread.Sleep(200);
                }
            }
            stopWatch.Stop();
            //Console.WriteLine();
            _DataLoggingList = DataLoggingList;

            //Console.WriteLine("Iterated {0} number of times", index);
            TestReadout();
            //Console.WriteLine();
        }
        private void TestReadout()
        {
            /*Console.WriteLine();
            Console.WriteLine("Gotten {0} datapoints during the last test", _DataLoggingList.Count());
            Console.WriteLine();*/
            foreach (TestSequenceElement snippet in _DataLoggingList)
            {
                Console.WriteLine("Voltage {0}, Current {1}, Power {2}, Resistance {3} was measured at time {4}", snippet.GetVoltage(), snippet.GetCurrent(), snippet.GetPower(), snippet.GetResistance(), snippet.GetElapsedTime());
                lbLoggedData.Items.Add(string.Format(LoggingLayout, "Voltage: " + snippet.GetVoltage().ToString() + "V", "Current: " + snippet.GetCurrent().ToString() + "A", "Power: " + snippet.GetPower().ToString() + "W", "Resistance: " + snippet.GetResistance().ToString() + "Ohms", "at time: " + snippet.GetElapsedTime().ToString() + "ms"));
            }
        }
        #endregion

        private void btnComPortRefresh_Click(object sender, EventArgs e)
        {
            ComPortInit();
        }
        #region ComPort initialisation
        private void ComPortInit()
        {
            cbbComPort.Text = ("");
            //Initialises the ComPorts available
            string[] ports = SerialPort.GetPortNames();
            Console.WriteLine("The following serial ports were found:");
            foreach (string port in ports)
            {
                Console.WriteLine(port);
                cbbComPort.Items.Add(port);
                cbbComPort.Text = (port.ToString());
            }
        }
        #endregion

        private void btnBeepState_Click(object sender, EventArgs e)
        {
            ToggleBeeper();
        }

        private void btnSaveLog_Click(object sender, EventArgs e)
        {
            SaveTestResults();
        }



        private void btnSaveTest_Click(object sender, EventArgs e)
        {
            SaveTestSequence();
        }

        private void btnLoadTest_Click(object sender, EventArgs e)
        {
            LoadTestSequence();
        }


        private void SaveTestSequence()
        {
            Console.WriteLine("Saving test sqeuence");
            FilePath = "";
            FileName = "";

            //If File path text box is empty
            if (string.IsNullOrWhiteSpace(txtFilePath.Text))
            {
                MessageBox.Show("File Path is empty", "File Path is empty", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                FilePath = txtFilePath.Text;

            }
            //If File Name text box is empty
            if (string.IsNullOrWhiteSpace(txtFileName.Text))
            {
                MessageBox.Show("File name is empty", "File name is empty", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                FileName = txtFileName.Text;
            }

            try
            {
                List<TestSequenceElement> TestSequenceElements = _TestSequenceElements;


                _XmlSerial = new XmlSerializer(typeof(List<TestSequenceElement>));

                FileStream XmlStream = new FileStream(FilePath + @"\" + FileName + ".Xml", FileMode.Create, FileAccess.Write);

                _XmlSerial.Serialize(XmlStream, TestSequenceElements);

                FilePath = "";
                FileName = "";
                XmlStream.Close();
            }
            catch
            {
                MessageBox.Show("Error writing file, check filepaths are ok");
            }

            
        }

        private void SaveTestResults()
        {
            FilePath = "";
            FileName = "";

            //If File path text box is empty
            if (string.IsNullOrWhiteSpace(txtLogFilePath.Text))
            {
                MessageBox.Show("File Path is empty", "File Path is empty", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                FilePath = txtLogFilePath.Text;
            }
            //If File Name text box is empty
            if (string.IsNullOrWhiteSpace(txtLogResultsName.Text))
            {
                MessageBox.Show("File name is empty", "File name is empty", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                FileName = txtLogResultsName.Text;
            }

            try
            {
                List<TestSequenceElement> DataLoggingList = _DataLoggingList;

                //FileStream stream = new FileStream(FilePath + @"\" + FileName + ".csv", FileMode.Create, FileAccess.Write);
                StreamWriter stream = new StreamWriter(FilePath + @"\" + FileName + ".csv");
                string msg = "Sample, Voltage, Current, Power, Resistance";
                stream.WriteLine(msg);

                for (int i = 0; i <= DataLoggingList.Count() - 1; i++)
                {
                    stream.Write(i + ',' + DataLoggingList[i].GetVoltage().ToString() + ',' + DataLoggingList[i].GetCurrent().ToString() + ',' + DataLoggingList[i].GetPower().ToString() + ',' + DataLoggingList[i].GetResistance().ToString() + "\n\r");
                }
                stream.Close();
            }
            catch
            {
                MessageBox.Show("Try checking the filepath and file name", "Error opening file", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }


            /*

            _XmlSerial = new XmlSerializer(typeof(List<TestSequenceElement>));

            FileStream XmlStream = new FileStream(FilePath + @"\" + FileName + ".Xml", FileMode.Create, FileAccess.Write);

            _XmlSerial.Serialize(XmlStream, DataLoggingList);
            FilePath = "";
            FileName = "";
            XmlStream.Close();
            */
        }

        private void LoadTestSequence()
        {
            //If File path text box is empty
            if (string.IsNullOrWhiteSpace(txtFilePath.Text))
            {
                MessageBox.Show("File Path is empty", "File Path is empty", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                FilePath = txtFilePath.Text;
            }
            //If File Name text box is empty
            if (string.IsNullOrWhiteSpace(txtFileName.Text))
            {
                MessageBox.Show("File name is empty", "File name is empty", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                FileName = txtFileName.Text;
            }

            List<TestSequenceElement> TestSequenceElements = new List<TestSequenceElement>();
            _XmlSerial = new XmlSerializer(typeof(List<TestSequenceElement>));
            try
            {
                FileStream XmlStream = new FileStream(FilePath + @"\" + FileName + ".Xml", FileMode.Open, FileAccess.Read);
                TestSequenceElements = (List<TestSequenceElement>)_XmlSerial.Deserialize(XmlStream);

                _TestSequenceElements = TestSequenceElements;
                foreach (TestSequenceElement Element in _TestSequenceElements)
                {
                    lbxTestSequence.Items.Add(string.Format(lvlvComponentTesterBox,Element.GetVoltage() + "V", Element.GetCurrent() + "A", Element.GetTime() + "ms"));
                }

                FilePath = "";
                XmlStream.Close();
            }
            catch
            {
                MessageBox.Show("Try checking the filepath and file name", "Error opening file", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            


        }


        private void ToggleBeeper()
        {

            foreach (SerialDevice device in _DeviceList)
            {
                if (device._deviceName == txtDeviceName.Text)
                {

                    if (_Beeper)
                    {
                        Console.WriteLine("Changing beep statis");
                        device.SendQueuePush("BEEP", "0");
                        txtBeep.Text = "Off";
                        txtBeep.Refresh();
                        _Beeper = false;
                    }
                    else
                    {
                        Console.WriteLine("Changing beep statis");
                        device.SendQueuePush("BEEP", "1");
                        txtBeep.Text = "On";
                        txtBeep.Refresh();
                        _Beeper = true;
                    }
                }
            }
        }
    }
}