using SerialCommunicationService.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.ServiceProcess;
using System.Threading;
using System.Timers;

namespace SerialCommunicationService
{
    public partial class SerialCommunicationService : ServiceBase
    {
        System.Timers.Timer timer = new System.Timers.Timer();
        string configFileName = ConfigurationManager.AppSettings["config_file_name"];
        string restartServiceBatchFileName = ConfigurationManager.AppSettings["restart_service_batch_file_name"];
        string startServiceBatchFileName = ConfigurationManager.AppSettings["start_service_batch_file_name"];
        string stopServiceBatchFileName = ConfigurationManager.AppSettings["stop_service_batch_file_name"];
        string installServiceBatchFileName = ConfigurationManager.AppSettings["install_service_batch_file_name"];
        SerialPort serialPort1;
        string _checkcommandResponse, _fwDateCommandResponse, _fwCommandResponse, _tndCommandResponse,
            _getValuesResponse, _getCellResponse, _serialPortReply, serviceName, moduleStatus;
        string[] availablePorts;
        string[] lines;
        private static string configFilePath, rootPath;

        public SerialCommunicationService()
        {
            InitializeComponent();
        }

        protected string GetServiceName()
        {
            // Calling System.ServiceProcess.ServiceBase::ServiceNamea allways returns
            // an empty string,
            // see https://connect.microsoft.com/VisualStudio/feedback/ViewFeedback.aspx?FeedbackID=387024

            // So we have to do some more work to find out our service name, this only works if
            // the process contains a single service, if there are more than one services hosted
            // in the process you will have to do something else

            int processId = System.Diagnostics.Process.GetCurrentProcess().Id;
            string query = "SELECT * FROM Win32_Service where ProcessId = " + processId;
            System.Management.ManagementObjectSearcher searcher =
                new System.Management.ManagementObjectSearcher(query);

            foreach (System.Management.ManagementObject queryObj in searcher.Get())
            {
                return queryObj["Name"].ToString();
            }

            throw new Exception("Can not get the ServiceName");
        }

        protected override void OnStart(string[] args)
        {
            lines = new string[4];
            serviceName = GetServiceName();
            DataLoggerControl.Write($"Service has been started at { DateTime.Now} ", serviceName);
            rootPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, serviceName);
            configFilePath = Path.Combine(rootPath, configFileName);
            lines = File.ReadAllLines(configFilePath);
            serialPort1 = new SerialPort();

            availablePorts = SerialPort.GetPortNames();
            rootPath = AppDomain.CurrentDomain.BaseDirectory;
            string serviceRootFolderName = Path.Combine(rootPath, serviceName);
            configFilePath = Path.Combine(serviceRootFolderName, configFileName);
            if (!Directory.Exists(serviceRootFolderName))
            {
                DataLoggerControl.Write($"Service folder does not exist: { serviceRootFolderName}", serviceName);
                Directory.CreateDirectory(serviceRootFolderName);
                DataLoggerControl.Write($"Service root folder has been created : { serviceRootFolderName}", serviceName);
            }
            else
            {
                DataLoggerControl.Write($"Service root folder is: { serviceRootFolderName}", serviceName);
            }
            if (!File.Exists(configFilePath))
            {
                DataLoggerControl.Write($"Service configuration file does not exist: { configFilePath}", serviceName);
                File.Create(configFilePath);
                DataLoggerControl.Write($"Service configuration file has been created: { configFilePath}", serviceName);
            }
            else
                DataLoggerControl.Write($"Service configuration file path is: { configFilePath}", serviceName);

            //DeleteOldFile(Path.Combine(serviceRootFolderName, "Logs"));

            if (InitiateService())
            {
                DataLoggerControl.Write($"Service parameters have been initiated correctly!", serviceName);
                if (SendingSerialNumCommand())
                {
                    DataLoggerControl.Write($"SER_NUM# Command is successful!", serviceName);
                    if (CreateModuleOnStartup())
                    {
                        DataLoggerControl.Write($"Module has been created at the start of the service!", serviceName);
                    }
                    else
                    {
                        DataLoggerControl.Write($"Failed to create module at the start of the service!", serviceName);
                    }
                }
            }
            else
            {
                DataLoggerControl.Write($"Service parameters could not be set successfully! Please refresh your service.", serviceName);
            }

            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            if (lines[2] != null && Convert.ToInt32(lines[2]) > 0)
                timer.Interval = Convert.ToInt32(lines[2]);
            else
                timer.Interval = 180000;
            timer.Enabled = true;
        }


        /// <summary>
        /// Creating module at the start of the service
        /// </summary>
        private bool CreateModuleOnStartup()
        {
            #region GET_TND# and GET_FW_DATE#
            bool isFW_DATE = false;
            bool isFW = false;
            bool isTND = false;
            bool isModuleCreated = false;

            int portResponse = 4;
            int portResponseCounter = 0;

            while (portResponse > 0)
            {
                try
                {
                    if (serialPort1.IsOpen)
                    {
                        string portOpened = $"Port is opened while creating module";
                        Thread.Sleep(4000);
                        serialPort1.Write("CHECK#");
                        Thread.Sleep(4000);
                        _checkcommandResponse = serialPort1.ReadTo("!");
                        if (_checkcommandResponse.Contains("OKAY"))
                        {
                            DataLoggerControl.WriteDataLogs($"CHECK Command is successfull on port {serialPort1.PortName}", serviceName);
                            DataLoggerControl.WriteDataLogs($"CHECK Command response is: {_checkcommandResponse}{Environment.NewLine}", serviceName);

                            DataLoggerControl.WriteDataLogs($"Command sent: GET_TND#", serviceName);
                            Thread.Sleep(4000);
                            serialPort1.Write("GET_TND#");
                            Thread.Sleep(4000);
                            _tndCommandResponse = serialPort1.ReadTo("!");
                            Thread.Sleep(2000);
                            if (_tndCommandResponse.Contains("ERROR") || _tndCommandResponse == "")
                            {
                                string getTNDError = $"GET_TND# ERROR : {_tndCommandResponse}{Environment.NewLine}";
                                DataLoggerControl.WriteDataLogs(getTNDError, serviceName);
                            }
                            else
                            {
                                isTND = true;
                                _tndCommandResponse = (!string.IsNullOrEmpty(_tndCommandResponse) && _tndCommandResponse.Length > 5) ? _tndCommandResponse.Substring(0, _tndCommandResponse.Length - 5) : _tndCommandResponse;
                                DataLoggerControl.WriteDataLogs($"Command response: {_tndCommandResponse}{Environment.NewLine}", serviceName);
                            }

                            Thread.Sleep(2000);
                            DataLoggerControl.WriteDataLogs($"Command sent: GET_FW_DATE#", serviceName);
                            Thread.Sleep(4000);
                            serialPort1.Write("GET_FW_DATE#");
                            Thread.Sleep(4000);
                            _fwDateCommandResponse = serialPort1.ReadTo("!");
                            Thread.Sleep(2000);
                            if (_fwDateCommandResponse.Contains("ERROR") || _fwDateCommandResponse == "")
                            {
                                string getFWDate = $"GET_FW_DATE# ERROR response: {_fwDateCommandResponse}{Environment.NewLine}";
                                DataLoggerControl.WriteDataLogs(getFWDate, serviceName);
                            }
                            else
                            {
                                isFW_DATE = true;
                                _fwDateCommandResponse = (!string.IsNullOrEmpty(_fwDateCommandResponse) && _fwDateCommandResponse.Length > 4) ? _fwDateCommandResponse.Substring(0, _fwDateCommandResponse.Length - 5) : _fwDateCommandResponse;
                                DataLoggerControl.WriteDataLogs($"Command response: {_fwDateCommandResponse}{Environment.NewLine}", serviceName);
                            }

                            Thread.Sleep(4000);
                            DataLoggerControl.WriteDataLogs($"Command sent: GET_FW#", serviceName);
                            Thread.Sleep(4000);
                            serialPort1.Write("GET_FW#");
                            Thread.Sleep(4000);
                            _fwCommandResponse = serialPort1.ReadTo("!");
                            Thread.Sleep(2000);
                            if (_fwCommandResponse.Contains("ERROR") || _fwCommandResponse == "")
                            {
                                string getFWError = $"GET_FW# ERROR: {_fwCommandResponse}{Environment.NewLine}";
                                DataLoggerControl.WriteDataLogs(getFWError, serviceName);
                            }
                            else
                            {
                                isFW = true;
                                _fwCommandResponse = (!string.IsNullOrEmpty(_fwCommandResponse) && _fwCommandResponse.Length > 4) ? _fwCommandResponse.Substring(0, _fwCommandResponse.Length - 5) : _fwCommandResponse;
                                DataLoggerControl.WriteDataLogs($"Command response: {_fwCommandResponse}{Environment.NewLine}", serviceName);
                            }
                            portResponse = 0; portResponseCounter--;
                        }
                        else if (_checkcommandResponse.Contains("ERROR"))
                        {
                            string checkCommandError = $"CHECK# Command error while creating module {_checkcommandResponse}{Environment.NewLine}";
                            DataLoggerControl.WriteDataLogs(checkCommandError, serviceName);
                            portResponse--; portResponseCounter++;
                        }
                    }
                    else
                    {
                        string portNotOpened = $"Serial port {serialPort1.PortName} is not opened while creating module!";
                        DataLoggerControl.Write(portNotOpened, serviceName);
                    }

                    //Module creation at the start of the service
                    if (isFW_DATE && isTND && isFW && portResponseCounter < 4 && portResponse == 0)
                    {
                        moduleStatus = ModuleStatus.Online.ToString();
                        ModuleObjects moduleObject = PrepareModuleObject(_fwDateCommandResponse, _tndCommandResponse,
                                                                         _fwCommandResponse, lines[3], lines[0],
                                                                         lines[1], moduleStatus, string.Empty);
                        AppService.RunModulePostAsync(moduleObject, serviceName);
                        isModuleCreated = true;
                    }
                    else
                        moduleStatus = ModuleStatus.Offline.ToString();

                    ErrorModuleObject errorModuleObject = new ErrorModuleObject();
                    errorModuleObject.serialNumber = _serialPortReply;
                    errorModuleObject.status = moduleStatus;
                    AppService.RunUpdateModuleStatusPostAsync(errorModuleObject, serviceName);
                }
                catch (Exception ex)
                {
                    DataLoggerControl.Write($"ERROR! {ex}", serviceName);
                    DataLoggerControl.WriteDataLogs($"{ex.Message}", serviceName);
                    _tndCommandResponse = string.Empty;
                    _fwDateCommandResponse = string.Empty;
                    _fwCommandResponse = string.Empty;
                    _checkcommandResponse = string.Empty;
                    portResponse--; portResponseCounter++; 
                }
            }
            _tndCommandResponse = string.Empty;
            _fwDateCommandResponse = string.Empty;
            _fwCommandResponse = string.Empty;
            _checkcommandResponse = string.Empty;

            return isModuleCreated;
            #endregion GET_TND# and GET_FW_DATE#
        }

        /// <summary>
        /// Sending serial command
        /// </summary>
        /// <returns></returns>
        private bool SendingSerialNumCommand()
        {
            #region GET_SER_NUM#
            bool isSerialNumCommandOk = false;
            if (serialPort1.IsOpen)
            {
                int portResponse = 3;
                while (portResponse > 0)
                {
                    try
                    {
                        Thread.Sleep(4000);
                        serialPort1.Write("CHECK#");
                        Thread.Sleep(4000);
                        _checkcommandResponse = serialPort1.ReadTo("!");
                        if (_checkcommandResponse.Contains("OKAY"))
                        {
                            DataLoggerControl.WriteDataLogs($"CHECK Command is successfull on port {serialPort1.PortName}", serviceName);
                            DataLoggerControl.WriteDataLogs($"CHECK Command response is: {_checkcommandResponse}{Environment.NewLine}", serviceName);
                            DataLoggerControl.WriteDataLogs($"Command sent: GET_SER_NUM#", serviceName);
                            Thread.Sleep(4000);
                            serialPort1.Write("GET_SER_NUM#");
                            Thread.Sleep(4000);
                            _serialPortReply = serialPort1.ReadTo("!");
                            Thread.Sleep(2000);
                            if (_serialPortReply.Contains("ERROR") || _serialPortReply == "")
                            {
                                DataLoggerControl.WriteDataLogs($"GET_SER_NUM# ERROR response: {_serialPortReply}{Environment.NewLine}", serviceName);
                            }
                            else
                            {
                                isSerialNumCommandOk = true;
                                _serialPortReply = (!string.IsNullOrEmpty(_serialPortReply) && _serialPortReply.Length > 4) ? _serialPortReply.Substring(0, _serialPortReply.Length - 4) : _serialPortReply;
                                DataLoggerControl.WriteDataLogs($"Command response: {_serialPortReply}{Environment.NewLine}", serviceName);
                            }
                            portResponse = 0;
                        }
                        else if (_checkcommandResponse.Contains("ERROR"))
                        {
                            DataLoggerControl.WriteDataLogs($"CHECK# Command Error: {_checkcommandResponse}{Environment.NewLine}", serviceName);
                            portResponse--;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataLoggerControl.Write($"ERROR! {ex}", serviceName);
                        DataLoggerControl.WriteDataLogs($"GET_SER_NUM# Command: {ex.Message}", serviceName);
                        _serialPortReply = string.Empty;
                        _checkcommandResponse = string.Empty;
                        portResponse--;
                    }
                }
            }

            return isSerialNumCommandOk;
            #endregion
        }

        /// <summary>
        /// Preparing Module Object
        /// </summary>
        /// <param name="_fwDateCommandResponse"></param>
        /// <param name="_tndCommandResponse"></param>
        /// <param name="_fwCommandResponse"></param>
        /// <param name="siteName"></param>
        /// <returns></returns>
        private ModuleObjects PrepareModuleObject(string _fwDateCommandResponse, string _tndCommandResponse,
                                                  string _fwCommandResponse, string siteName, string comPort,
                                                  string boudRate, string status, string capacity)
        {
            if (string.IsNullOrEmpty(_fwDateCommandResponse))
            {
                throw new ArgumentException($"'{nameof(_fwDateCommandResponse)}' cannot be null or empty", nameof(_fwDateCommandResponse));
            }

            if (string.IsNullOrEmpty(_tndCommandResponse))
            {
                throw new ArgumentException($"'{nameof(_tndCommandResponse)}' cannot be null or empty", nameof(_tndCommandResponse));
            }

            if (_fwCommandResponse is null)
            {
                throw new ArgumentNullException(nameof(_fwCommandResponse));
            }

            if (string.IsNullOrEmpty(siteName))
            {
                throw new ArgumentException($"'{nameof(siteName)}' cannot be null or empty", nameof(siteName));
            }

            if (string.IsNullOrEmpty(comPort))
            {
                throw new ArgumentException($"'{nameof(comPort)}' cannot be null or empty", nameof(comPort));
            }

            if (string.IsNullOrEmpty(boudRate))
            {
                throw new ArgumentException($"'{nameof(boudRate)}' cannot be null or empty", nameof(boudRate));
            }

            if (string.IsNullOrEmpty(status))
            {
                throw new ArgumentException($"'{nameof(status)}' cannot be null or empty", nameof(status));
            }

            ModuleObjects moduleObject = new ModuleObjects();
            moduleObject.serialNumber = _serialPortReply;
            moduleObject.name = _serialPortReply;
            moduleObject.sitesName = siteName;
            moduleObject.geT_FW_DATE = _fwDateCommandResponse;
            moduleObject.geT_TND = _tndCommandResponse;
            moduleObject.geT_FW = _fwCommandResponse;
            moduleObject.comport = comPort;
            moduleObject.boudrate = boudRate;
            moduleObject.capacity = capacity;
            moduleObject.status = status;
            DataLoggerControl.Write($"Module object prepared: {moduleObject.serialNumber}, {moduleObject.name}, {moduleObject.sitesName}," +
                $" {moduleObject.geT_FW_DATE}, {moduleObject.geT_TND}, {moduleObject.geT_FW}, {moduleObject.comport}, {moduleObject.boudrate}," +
                $" {moduleObject.status}, {moduleObject.capacity} ", serviceName);
            return moduleObject;
        }

        /// <summary>
        /// Creating test module 
        /// </summary>
        public bool CreateTestModule()
        {
            bool isTestModuleCreated = false;

            try
            {
                List<string> getValuesAndGetCells = new List<string>();
                bool isGetValues = false;
                bool isGetCell = false;
                if (serialPort1.IsOpen)
                {
                    string portOpened = $"CreateTestModule: COM port is open and starting reading data from {serialPort1.PortName}";
                    DataLoggerControl.Write(portOpened, serviceName);

                    #region GET_VALUES# and GET_CELL#
                    int portResponse = 4;
                    int portResponseCounter = 0;
                    while (portResponse > 0)
                    {
                        try
                        {
                            Thread.Sleep(4000);
                            serialPort1.Write("CHECK#");
                            Thread.Sleep(4000);
                            _checkcommandResponse = serialPort1.ReadTo("!");
                            if (_checkcommandResponse.Contains("OKAY"))
                            {
                                DataLoggerControl.WriteDataLogs($"CHECK Command is successfull on port {serialPort1.PortName}", serviceName);
                                DataLoggerControl.WriteDataLogs($"CHECK Command response is: {_checkcommandResponse}{Environment.NewLine}", serviceName);

                                DataLoggerControl.WriteDataLogs($"Command sent: GET_VALUES#", serviceName);
                                Thread.Sleep(4000);
                                serialPort1.Write("GET_VALUES#");
                                Thread.Sleep(4000);
                                _getValuesResponse = serialPort1.ReadTo("!");
                                Thread.Sleep(2000);
                                if (_getValuesResponse.Contains("ERROR") || _getValuesResponse == "")
                                {
                                    string getValuesError = $"GET_VALUES# ERROR response: {_getValuesResponse}{Environment.NewLine}";
                                    DataLoggerControl.WriteDataLogs(getValuesError, serviceName);
                                }
                                else
                                {
                                    _getValuesResponse = (!string.IsNullOrEmpty(_getValuesResponse) && _getValuesResponse.Length > 5) ? _getValuesResponse.Substring(0, _getValuesResponse.Length - 5) : _getValuesResponse;
                                    getValuesAndGetCells.Add(_getValuesResponse);
                                    isGetValues = true;
                                    DataLoggerControl.WriteDataLogs($"Command response: {_getValuesResponse}{Environment.NewLine}", serviceName);
                                }

                                Thread.Sleep(4000);
                                DataLoggerControl.WriteDataLogs($"Command sent: GET_CELL#", serviceName);
                                Thread.Sleep(4000);
                                serialPort1.Write("GET_CELL#");
                                Thread.Sleep(4000);
                                _getCellResponse = serialPort1.ReadTo("!");
                                Thread.Sleep(2000);
                                if (_getCellResponse.Contains("ERROR") || _getCellResponse == "")
                                {
                                    string getCellError = $"GET_CELL# ERROR: {_getCellResponse}{Environment.NewLine}";
                                    DataLoggerControl.WriteDataLogs(getCellError, serviceName);
                                }
                                else
                                {
                                    _getCellResponse = (!string.IsNullOrEmpty(_getCellResponse) && _getCellResponse.Length > 5) ? _getCellResponse.Substring(0, _getCellResponse.Length - 5) : _getCellResponse;
                                    getValuesAndGetCells.Add(_getCellResponse);
                                    isGetCell = true;
                                    DataLoggerControl.WriteDataLogs($"Command response: {_getCellResponse}{Environment.NewLine}", serviceName);
                                }
                                portResponse = 0; //portResponseCounter--;
                            }
                            else if (_checkcommandResponse.Contains("ERROR"))
                            {
                                string checkCommandError = $"CHECK# Command error while creating test module {_checkcommandResponse}{Environment.NewLine}";
                                DataLoggerControl.WriteDataLogs(checkCommandError, serviceName);
                                portResponse--; //portResponseCounter++;
                            }
                        }
                        catch (Exception ex)
                        {
                            DataLoggerControl.Write($"ERROR! {ex}", serviceName);
                            DataLoggerControl.WriteDataLogs($"GET_VALUES# Command: {ex.Message}", serviceName);
                            _getValuesResponse = string.Empty;
                            _getCellResponse = string.Empty;
                            _checkcommandResponse = string.Empty;
                            portResponse--; //portResponseCounter++;
                        }
                    }

                    //if (isGetCell && isGetValues && portResponseCounter < 4 && portResponse == 0)
                    if (isGetCell && isGetValues && portResponseCounter < 4 && portResponse == 0)
                    {
                        //moduleStatus = ModuleStatus.Online.ToString();
                        foreach (var item in getValuesAndGetCells)
                        {
                            TestModuleObject testModuleObj = new TestModuleObject();
                            testModuleObj.serialNumber = _serialPortReply;
                            testModuleObj.inputValues = item;
                            testModuleObj.id = 0;
                            AppService.RunCreateValuesAndCellsPostAsync(testModuleObj, serviceName);
                            isTestModuleCreated = true;
                        }
                    }
                    //else
                    //    moduleStatus = ModuleStatus.Offline.ToString();

                    //AppService.RunUpdateModuleStatusPostAsync(moduleStatus, _serialPortReply, serviceName);
                    #endregion GET_VALUES# and GET_CELL#
                }
                else
                {
                    string portnotOpened = $"COM Port {serialPort1.PortName} is not opened while creating test module!";
                    DataLoggerControl.Write(portnotOpened, serviceName);
                }
                _getValuesResponse = string.Empty;
                _checkcommandResponse = string.Empty;
                _getCellResponse = string.Empty;
            }
            catch (Exception ex)
            {
                _getValuesResponse = string.Empty;
                _checkcommandResponse = string.Empty;
                _getCellResponse = string.Empty;
                DataLoggerControl.Write($"ERROR! {ex}", serviceName);
            }

            return isTestModuleCreated;
        }

        public static void DeleteOldFile(string directoryPath)
        {
            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);
            string destinationDirectory = Path.Combine(directoryPath, "Logs");
            string[] files = Directory.GetFiles(destinationDirectory);
            if (!Directory.Exists(destinationDirectory))
                Directory.CreateDirectory(destinationDirectory);
            foreach (string file in files)
            {
                FileInfo fi = new FileInfo(file);
                if (fi.LastAccessTime < DateTime.Now.AddDays(-2))
                {
                    File.Delete(file);
                }
            }
        }

        protected override void OnStop()
        {
            DataLoggerControl.Write($"Service is stopped at { DateTime.Now}", serviceName);

            if (serialPort1.IsOpen)
                serialPort1.Close();
        }
        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            if (CreateTestModule())
                DataLoggerControl.Write($"Test module has been created!", serviceName);
        }

        private bool InitiateService()
        {
            bool isServiceInitiated = false;

            try
            {
                availablePorts = SerialPort.GetPortNames();
                if (!File.Exists(configFilePath))
                {
                    File.Create(configFilePath);
                }

                if (availablePorts.Length > 0)
                {
                    if (File.Exists(configFilePath))
                    {
                        lines = File.ReadAllLines(configFilePath);

                        if (serialPort1.IsOpen)
                        {
                            serialPort1.Close();
                        }
                        serialPort1.PortName = lines[0];
                        serialPort1.Open();
                        if (availablePorts.Contains(lines[0]))
                        {
                            serialPort1.BaudRate = Convert.ToInt32(lines[1]);
                            serialPort1.ReadTimeout = 120000;
                            isServiceInitiated = true;
                        }
                        else
                        {
                            DataLoggerControl.Write($"Please update your configurations from available ports and their respective baud rates", serviceName);
                            serialPort1.Close();
                        }
                    }
                }
                else
                {
                    DataLoggerControl.Write($"COM Port is not available. Please check your ports and try again.", serviceName);
                }
            }
            catch (Exception ex)
            {
                DataLoggerControl.Write($"ERROR! Initialising Reading Error {ex}", serviceName);
                serialPort1.Close();
                isServiceInitiated = false;
            }

            return isServiceInitiated;
        }
    }
}
