using AndreasReitberger.SCPI.Enums;
using CommunityToolkit.Mvvm.ComponentModel;
using System.Globalization;
using System.IO.Ports;

namespace AndreasReitberger.SCPI
{
    public partial class SCPIClient : ObservableObject
    {
        #region Fields

        private static SerialPort? _port;
        #endregion

        #region Properties

        public static bool IsInitialized => _port?.IsOpen == true;

        #endregion
        #region Public Static Methods
        /// <summary>  Availables serialPortName ports.</summary>
        /// <returns>Returns all available serialPortName Ports</returns>
        public static string[] GetAvailablePorts() => SerialPort.GetPortNames();
        
        /// <summary>Gets the available outputs for the Toellner TOE 8952</summary>
        /// <returns>An array of the available Outputs</returns>
        public static string[] GetOutputs() => Enum.GetNames(typeof(Output));

        #endregion

        #region Public Methods
        /// <summary>Initializes the the SerialPort ant the connection to the Toellner PSU.</summary>
        /// <param name="serialPortName">The serialPortName-Port (for instance "COM1").</param>
        /// <param name="baudrate">The baudrate (for instance 9600).</param>
        /// <returns>True on success</returns>
        public bool Init(string serialPortName, int baudrate)
        {
            try
            {
                if (_port?.IsOpen is true)
                    _port.Close();
                _port = new SerialPort(serialPortName, baudrate,
                    Parity.None, 8, StopBits.One
                    );
                _port.Open();
                //:SYST:REM
                if (_port?.IsOpen is true)
                {
                    SerialWriteRead($":{SCPICommand.SYST}:{SCPICommand.REM}\n");
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>Initializes the the SerialPort ant the connection to the Toellner PSU.</summary>
        /// <param name="serialPortName">The serialPortName-Port (for instance "COM1").</param>
        /// <param name="baudrate">The baudrate (for instance 9600).</param>
        /// <param name="parity">The parity bit(s).</param>
        /// <param name="dataBits">The data bits.</param>
        /// <param name="stopBits">The stop bit(s).</param>
        /// <returns>True on success</returns>
        public bool Init(string serialPortName, int baudrate, Parity parity, int dataBits, StopBits stopBits)
        {
            try
            {
                if (_port?.IsOpen is true)
                    _port.Close();
                _port = new SerialPort(serialPortName, baudrate,
                    parity, dataBits, stopBits
                    );
                _port.Open();
                //:SYST:REM
                if (_port?.IsOpen is true)
                {
                    SerialWriteRead($":{SCPICommand.SYST}:{SCPICommand.REM}\n");
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>Sets the output level for the specified output.</summary>
        /// <param name="output">The output.</param>
        /// <param name="voltage">The voltage.</param>
        /// <param name="turnOutputOn">if set to <c>true</c> [turn output on].</param>
        /// <returns>True on success</returns>
        public bool SetOutputLevel(Output output, double voltage, bool turnOutputOn = false)
        {
            //:INST OUT1;:VOLT 10;OUTP 1
            if (_port?.IsOpen is true)
            {
                SerialWriteRead(
                    $":{SCPICommand.INST} {output};" +
                    $":{SCPICommand.VOLT} {voltage}; " +
                    $"{SCPICommand.OUTP} {(turnOutputOn ? "1" : "0")}"
                    );
                return true;
            }
            else return false;
        }

        /// <summary>Sets the output level on the specified output.</summary>
        /// <param name="outputName">Name of the output.</param>
        /// <param name="voltage">The voltage.</param>
        /// <param name="turnOutputOn">if set to <c>true</c> [turn output on].</param>
        /// <returns>True on success</returns>
        public bool SetOutputLevel(string outputName, double voltage, bool turnOutputOn = false)
        {
            //:INST OUT1;:VOLT 10;OUTP 1
            if (_port?.IsOpen is true)
            {
                SerialWriteRead(
                    $":{SCPICommand.INST} {outputName};" +
                    $":{SCPICommand.VOLT} {voltage}; " +
                    $"{SCPICommand.OUTP} {(turnOutputOn ? "1" : "0")}"
                    );
                return true;
            }
            else return false;
        }

        /// <summary>Sets the output level the specified output.</summary>
        /// <param name="output">The output.</param>
        /// <param name="voltage">The voltage.</param>
        /// <param name="current">The current.</param>
        /// <param name="turnOutputOn">if set to <c>true</c> [turn output on].</param>
        /// <returns>True on success</returns>
        public bool SetOutputLevel(Output output, double voltage, double current, bool turnOutputOn = false)
        {
            //:INST OUT1;:VOLT 10;CURR 1;OUTP 1
            if (_port?.IsOpen is true)
            {
                SerialWriteRead(
                    $":{SCPICommand.INST} {output};" +
                    $":{SCPICommand.VOLT} {voltage.ToString("0.00", CultureInfo.InvariantCulture)};" +
                    $":{SCPICommand.CURR} {current.ToString("0.00", CultureInfo.InvariantCulture)}; " +
                    $"{SCPICommand.OUTP} {(turnOutputOn ? "1" : "0")}"
                    );
                return true;
            }
            else return false;
        }

        /// <summary>Sets the output level on the specified output.</summary>
        /// <param name="outputName">Name of the output.</param>
        /// <param name="voltage">The voltage.</param>
        /// <param name="current">The current.</param>
        /// <param name="turnOutputOn">if set to <c>true</c> [turn output on].</param>
        /// <returns>True on success</returns>
        public bool SetOutputLevel(string outputName, double voltage, double current, bool turnOutputOn = false)
        {
            //:INST OUT1;:VOLT 10;CURR 1;OUTP 1
            if (_port?.IsOpen is true)
            {
                SerialWriteRead(
                    $":{ SCPICommand.INST} {outputName};" +
                    $":{SCPICommand.VOLT} {voltage};" +
                    $":{SCPICommand.CURR} {current}; " +
                    $"{SCPICommand.OUTP} {(turnOutputOn ? "1" : "0")}"
                    );
                return true;
            }
            else return false;
        }

        /// <summary>Gets the current value from the passed output.</summary>
        /// <param name="output">The output (0 &lt;&gt; 1).</param>
        /// <param name="target">The target value (voltage &lt;&gt; current &lt;&gt; Power).</param>
        /// <returns>The value as double</returns>
        public double GetValue(Output output, ValueTarget target)
        {
            //:INST OUT1;:MEAS:VOLT?;:MEAS:CURR?;:MEAS:POW?;:VOLT:PROT:STAT?
            if (_port?.IsOpen is true)
            {
                string result = 
                    SerialWriteRead(
                        $":{SCPICommand.INST} {output};" +
                        $":{SCPICommand.MEAS}:{SCPICommand.VOLT}?;" +
                        $":{SCPICommand.MEAS}:{SCPICommand.CURR}?;" +
                        $":{SCPICommand.MEAS}:{SCPICommand.POW}?;" +
                        $":{SCPICommand.VOLT}:{SCPICommand.PROT}:{SCPICommand.STAT}?\n");
                double res = 0;
                if (!string.IsNullOrEmpty(result))
                {
                    try
                    {
                        string[] parts = result.Split(';');
                        //res = Convert.ToDouble(parts[0]);
                        res = double.Parse(parts[(int)target], CultureInfo.InvariantCulture);
                    }
                    catch (Exception)
                    {
                        return -1;
                    }
                }
                return res;
            }
            else return 0;
        }

        /// <summary>Toggles the output on or off.</summary>
        /// <param name="state">  True for on, false for off</param>
        /// <returns>True on success</returns>
        public bool ToggleOutput(bool state)
        {
            //;OUTP 1|0
            if (_port?.IsOpen is true)
            {
                SerialWriteRead($"{SCPICommand.OUTP} {(state ? "1" : "0")}");
                return true;
            }
            else return false;
        }


        /// <summary>  Clears the Power Supply levels</summary>
        /// <returns>True on success</returns>
        public bool RemoteClear()
        {
            if (_port?.IsOpen is true)
            {
                SerialWriteRead($"*{SCPICommand.RST};*{SCPICommand.CLS};");
                return true;
            }
            else return false;
        }
        /// <summary>Sends a plain command string.</summary>
        /// <param name="command">The command.</param>
        public void SendCommand(string command)
        {
            if (_port?.IsOpen is true)
            {
                SerialWriteRead(command);
            }
        }

        /// <summary>Closes the serial connection to the DUT.</summary>
        /// <returns>True on succeeded</returns>
        public bool Close()
        {
            if (_port?.IsOpen is true)
            {
                //:SYST:LOC
                SerialWriteRead($":{SCPICommand.SYST}:{SCPICommand.LOC}\n");
                _port.Close();
                return true;
            }
            else
                return false;
        }

        #endregion

        #region Private methods
        /// <summary>  Sends a command via the serial port to the DUT</summary>
        /// <param name="command">The command.</param>
        /// <returns>The result sended by the DUT</returns>
        private static string SerialWriteRead(string command)
        {
            string result = string.Empty;
            if (_port != null && _port.IsOpen)
            {
                //_port.Encoding = Encoding.GetEncoding(1252);
                _port.WriteLine(command);
                Thread.Sleep(500);
                if (_port.BytesToRead > 0)
                {
                    //result = cleanReturnString(_port.ReadExisting(), command);
                    result = _port.ReadExisting();
                }
            }
            return result;
        }
        #endregion

    }
}
