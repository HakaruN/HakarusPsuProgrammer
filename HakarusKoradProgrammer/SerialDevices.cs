using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO.Ports;
using System.Windows.Forms;

namespace HakarusKoradProgrammer
{
    class SerialDevice
    {
        //public variables
        public string _comPort = "";
        public string _deviceName = "";
        public bool _isConnected = false;
        public bool _isOn = false;
        public string _IDN = "";
        
        //maybe unused idk
        public string _voltage = "0";
        public string _current = "0";

        public float _fvoltage = 0;
        public float _fcurrent = 0;

        //Connection settings
        private int _baudRate = 0;
        private int _dataBits = 0;
        private string _parity = "0";
        private string _stopBits = "0";
        private string _flowControl = "0";

        private string _ConsoleReturn = "";
        private string _returnBuffer = "";
        //private ushort _returnByte;

        public  int _returnDelay = 50;
        Thread PowerPolling;
        Thread DispatchQueueOperator;

        struct SerialMsg
        {
            string _instruction;
            string _value;

            public SerialMsg(string instruction, string value)
            {
                _instruction = instruction;
                _value = value;
            }
            public string GetInstruction()
            {
                return _instruction;
            }
            public string GetValue()
            {
                return _value;
            }
        }

        //Thread control flags for the devices serial data dispatch queueing thread
        private bool _ThreadEnd = false;
        private bool _dispatchQueuePause = false;

        SerialPort _port;

        Queue<SerialMsg> _dispatchQueue = new Queue<SerialMsg>();

        public SerialDevice(String comPort, int baudRate, int dataBits, string parity, string stopBits, string flowControl, string deviceName)
        {
            _isConnected = false;

            _port = new SerialPort();
            _deviceName = deviceName;
            _comPort = comPort;

            _baudRate = baudRate;
            _dataBits = dataBits;
            _parity = parity;
            _stopBits = stopBits;
            _flowControl = flowControl;

            _voltage = "0.0";
            _current = "0.0";

            //This spawns a thread that will poll the currently selected device and set the actual volt/current output txtbox
            _ThreadEnd = false;
            _dispatchQueuePause = false;

            DispatchQueueOperator = new Thread(DispatchQueueOperation);
            Console.WriteLine("Dispatch operator thread created");
            DispatchQueueOperator.IsBackground = true;
            DispatchQueueOperator.Start();

            PowerPolling = new Thread(PowerPoller);
            Console.WriteLine("polling thread created");
            PowerPolling.IsBackground = true;
            PowerPolling.Start();

            

            //Need to take in the COMport, BaudRate, DataBits, Parity, StopBits, FlowControl
        }

        ~SerialDevice()
        {
            _ThreadEnd = true;
            PowerPolling.Abort();
            if(PowerPolling.IsAlive)
            {
                Console.WriteLine("Power polling thread not terminated");
            }
            else
            {
                Console.WriteLine("Power Polling thread terminated");
            }
        }

        public bool Connect()
        {

            if (!_isConnected)
            {
                

                Console.WriteLine("Starting");
                _port.PortName = _comPort;
                Console.WriteLine("Port set to: {0}", _comPort);
                _port.BaudRate = _baudRate;
                Console.WriteLine("Baudrate set to: {0}", _baudRate);
                _port.DataBits = _dataBits;
                Console.WriteLine("Transfer width set to: {0}", _dataBits);

                if (_parity == "None")
                {
                    _port.Parity = Parity.None;
                    Console.WriteLine("Parity set to: None");
                }
                else if (_parity == "Odd")
                {
                    _port.Parity = Parity.Odd;
                    Console.WriteLine("Parity set to: Odd");
                }
                else if (_parity == "Even")
                {
                    _port.Parity = Parity.Even;
                    Console.WriteLine("Parity set to: Even");
                }
                else if (_parity == "Mark")
                {
                    _port.Parity = Parity.Mark;
                    Console.WriteLine("Parity set to: Mark");
                }
                else if (_parity == "Space")
                {
                    _port.Parity = Parity.Space;
                    Console.WriteLine("Parity set to: Space");
                }
                else
                { Console.WriteLine("Error setting parity"); return false; }

                if (_stopBits == "One")
                {
                    _port.StopBits = StopBits.One;
                    Console.WriteLine("Stop bits set to: One");
                }
                else if (_stopBits == "OnePointFive")
                {
                    _port.StopBits = StopBits.OnePointFive;
                    Console.WriteLine("Stop bits set to: One point Five");
                }
                else if (_stopBits == "Two")
                {
                    _port.StopBits = StopBits.Two;
                    Console.WriteLine("Stop bits set to: Two");
                }
                else
                {
                    { Console.WriteLine("Error setting Stop bits"); return false; }
                }

                if (_flowControl == "None")
                {
                    _port.Handshake = Handshake.None;
                    Console.WriteLine("Flow Control set to: None");
                }
                else if (_flowControl == "RequestToSend")
                {
                    _port.Handshake = Handshake.RequestToSend;
                    Console.WriteLine("Flow Control set to: RequestToSend");
                }
                else if (_flowControl == "RequestToSendXOnXOff")
                {
                    _port.Handshake = Handshake.RequestToSendXOnXOff;
                    Console.WriteLine("Flow Control set to: RequestToSendXOnXOff");
                }
                else if (_flowControl == "XOnXOff")
                {
                    _port.Handshake = Handshake.XOnXOff;
                    Console.WriteLine("Flow Control set to: XOnXOff");
                }
                else
                {
                    { Console.WriteLine("Error setting Stop bits"); return false; }
                }

                try { _port.Open();
                    _isConnected = true;
                    return true;
                }
                catch
                {
                    MessageBox.Show("Error Opening port", "Error Opening portr", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    _isConnected = false;
                    return false;
                }


            }
            else
            {
                MessageBox.Show("Error: serial device alread connected", "Still Connected", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _isConnected = false;
                return false;
            }
            
        }

        public bool Disconnect()
        {
            _port.Close();
            _isConnected = false;
            _ThreadEnd = true;
            PowerPolling.Abort();
            Console.WriteLine("Port closed.");
            Console.WriteLine();
            return false;
        }

        private void SendData(SerialMsg Packet)
        {
            string msg = "";

            string instruction = Packet.GetInstruction();
            string value = Packet.GetValue();

            

            msg = instruction + value;
            switch (instruction)
            {

                case "BEEP":
                    _returnBuffer = "";
                    //Console.WriteLine("Turning devicice beeper : {0}", value);
                    DataGetter(msg);
                    Thread.Sleep(_returnDelay);
                    break;
                case "OUT":                    
                    _returnBuffer = "";
                    //Console.WriteLine("Turning device power to: {0}", value);
                    DataGetter(msg);
                    Thread.Sleep(_returnDelay);
                    break;

                case "VSET1:":                    
                    _returnBuffer = "";
                    //Console.WriteLine("Sending voltage value: {0}",value);
                    _port.Write(msg);
                    //Console.WriteLine("Data sent: {0}", msg);
                    Thread.Sleep(_returnDelay);
                    break;

                case "ISET1:":
                    _returnBuffer = "";
                    //Console.WriteLine("Sending current value: {0}",value);
                    _port.Write(msg);
                    //Console.WriteLine("Data sent: {0}", msg);
                    Thread.Sleep(_returnDelay);
                    break;

                case "OCP":                    
                    _returnBuffer = "";
                    //Console.WriteLine("Setting OCP to: {0}", value);
                    _port.Write(msg);
                    Thread.Sleep(_returnDelay);
                    break;

                case "OVP":
                    _returnBuffer = "";
                    //Console.WriteLine("Setting OVP to: {0}", value);
                    _port.Write(msg);
                    Thread.Sleep(_returnDelay);
                    break;

                case "RCL":
                    _returnBuffer = "";
                    //Console.WriteLine("recalling state from Mem loc{0}", value);
                    _port.Write(msg);
                    Thread.Sleep(_returnDelay);
                    break;

                case "SAV":
                    _returnBuffer = "";
                    //Console.WriteLine("Saving state to Mem loc{0}",value);
                    _port.Write(msg);
                    Thread.Sleep(_returnDelay);
                    break;

                case "VOUT1?":
                    _returnBuffer = "";
                    DataGetter(msg);
                    Thread.Sleep(_returnDelay);
                    try { _fvoltage = float.Parse(_returnBuffer); }
                    catch { Console.WriteLine("Voltage sense error"); }
                    break;

                case "IOUT1?":
                    _returnBuffer = "";
                    DataGetter(msg);
                    Thread.Sleep(_returnDelay);
                    try { _fcurrent = float.Parse(_returnBuffer); }
                    catch { Console.WriteLine("Current sense error"); }
                    break;

                case "*IDN?":
                    _returnBuffer = "";
                    //Console.WriteLine("Getting device decription");
                    DataGetter(msg);
                    Thread.Sleep(_returnDelay);
                    _IDN = _returnBuffer;
                    //Console.Write("Connected: ", _returnBuffer);
                    break;
                case "STATUS?":
                    _returnBuffer = "";
                    //Console.WriteLine("Getting device decription");
                    DataGetter(msg);
                    Thread.Sleep(_returnDelay);
                    _IDN = _returnBuffer;
                    //Console.Write("Connected: ", _returnBuffer);
                    break;

                case "CSL":
                    _ConsoleReturn = "";
                    _returnBuffer = "";
                    //Console.WriteLine("Sending console message");
                    DataGetter(value);
                    Thread.Sleep(_returnDelay);
                    _ConsoleReturn = _returnBuffer;
                    //Console.Write("Device returned: ", _ConsoleReturn);
                    //Console.WriteLine("(Output in Hex): ", String.Format("{0,10:X}", _ConsoleReturn));
                    break;

                default:
                    Console.WriteLine("Error, unsupported instruction");
                    _returnBuffer = "";
                    break;
            }
        }




        #region PowerSwitch
        public bool PowerSwitch()
        {
            if (_isOn)
            {
                SendQueuePush("OUT", "0");
                _isOn = false;
                return false;
            }
            else
            {
                SendQueuePush("OUT", "1");
                _isOn = true;
                return true;
            }
        }
        public void TurnOff()
        {
            SendQueuePush("OUT", "0");
            _isOn = false;
        }

        public void TurnOn()
        {
            SendQueuePush("OUT", "1");
            _isOn = true;
        }
        #endregion

        public void DataGetter(string msg)
        {
            if(_isConnected)
            {
                _port.Write(msg);
                _port.DataReceived += new SerialDataReceivedEventHandler(serialPort1_DataReceived);
            }

        }

        private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            byte[] buffer = new byte[1];
            buffer = Encoding.UTF8.GetBytes(_port.ReadExisting());
            _returnBuffer = _returnBuffer + System.Text.Encoding.Default.GetString(buffer);
        }




        #region Dispatch Queue 
        public void SendQueuePush(string msg, string value)
        {
            SerialMsg Packet = new SerialMsg(msg,value);
            _dispatchQueue.Enqueue(Packet);
        }
        #endregion

        private void PowerPoller()
        {
            while (!_ThreadEnd)
            {
                if(!_dispatchQueuePause)
                {
                    PowerSense();
                }
            }
        }

        private void DispatchQueueOperation()
        {
            SerialMsg msg;
            while (!_ThreadEnd)
            {
                if (_dispatchQueue.Count != 0)
                {
                    //Dequeue
                    Thread.Sleep(20);
                    msg = _dispatchQueue.Dequeue();
                    SendData(msg);
                }
            }
        }

        private void PowerSense()
        {
            if (_isConnected)
            {
                SendQueuePush("VOUT1?","");
                Thread.Sleep(40);
                SendQueuePush("IOUT1?","");
                Thread.Sleep(180);
                //The sleeps in here assure that the device has a 4hz Volt and current update.
                //The 30ms sleep between is to prevent two calls from being writen together.
                //Note this shouldn't happen anyway but you cant be too sure.
            }
        }
        
        public string ReturnConsole()
        {
            return _ConsoleReturn;
        }
    }
}
