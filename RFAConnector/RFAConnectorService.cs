using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RFAConnector
{
    public partial class RFAConnectorService : ServiceBase
    {
        private EventLog _eventLog1;
        private TcpClient _tcpClient;
        private CancellationTokenSource _cancellationTokenSource;
        private const int BUFFER_SIZE = 1024;

        public enum ServiceState
        {
            SERVICE_STOPPED = 0x00000001,
            SERVICE_START_PENDING = 0x00000002,
            SERVICE_STOP_PENDING = 0x00000003,
            SERVICE_RUNNING = 0x00000004,
            SERVICE_CONTINUE_PENDING = 0x00000005,
            SERVICE_PAUSE_PENDING = 0x00000006,
            SERVICE_PAUSED = 0x00000007,
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct ServiceStatus
        {
            public int dwServiceType;
            public ServiceState dwCurrentState;
            public int dwControlsAccepted;
            public int dwWin32ExitCode;
            public int dwServiceSpecificExitCode;
            public int dwCheckPoint;
            public int dwWaitHint;
        };

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(System.IntPtr handle, ref ServiceStatus serviceStatus);

        public RFAConnectorService()
        {
            InitializeComponent();
            _eventLog1 = new EventLog();
            _eventLog1.Source = "RFAConnectorSource";
            _eventLog1.Log = "RFAConnectorLog";
        }

        protected override void OnStart(string[] args)
        {
            _eventLog1.WriteEntry("Starting the RFAConnector Service.");

            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            _cancellationTokenSource = new CancellationTokenSource();

            bool _enableTCPConnection = bool.Parse(GetConfigValue("ENABLE_TCP_CONNECTION"));

            if (_enableTCPConnection)
            {
                _eventLog1.WriteEntry("Starting the TcpClient.");
                StartTcpClientAsync(_cancellationTokenSource.Token).ConfigureAwait(false);
            }
            else
            {
                _eventLog1.WriteEntry("TcpClient disabled, starting file watcher.");
                StartFileSystemWatcherAsync(_cancellationTokenSource.Token).ConfigureAwait(false);
            }

            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
        }

        protected override void OnStop()
        {
            _eventLog1.WriteEntry("Stopping the RFAConnector Service.");

            base.OnStop();

            _cancellationTokenSource?.Cancel();
            _tcpClient?.Close();

            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_STOP_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            serviceStatus.dwCurrentState = ServiceState.SERVICE_STOPPED;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
        }

        /// <summary>
        /// Initiates a TCP client connection to the configured host and port.
        /// Continuously listens for incoming data and processes it.
        /// Retries the connection in case of failures.
        /// </summary>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        private async Task StartTcpClientAsync(CancellationToken cancellationToken)
        {
            try
            {
                string host = GetConfigValue("RFA_TCP_HOST");
                int port = int.Parse(GetConfigValue("RFA_TCP_PORT"));

                while (!cancellationToken.IsCancellationRequested)
                {
                    try
                    {
                        using (_tcpClient = new TcpClient())
                        {
                            await _tcpClient.ConnectAsync(host, port);
                            _eventLog1.WriteEntry($"Connected to {host}:{port}");

                            using (NetworkStream stream = _tcpClient.GetStream())
                            {
                                byte[] buffer = new byte[BUFFER_SIZE];
                                while (!cancellationToken.IsCancellationRequested)
                                {
                                    try
                                    {
                                        int bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken);
                                        if (bytesRead == 0) break;

                                        string receivedData = Encoding.ASCII.GetString(buffer, 0, bytesRead);
                                        _eventLog1.WriteEntry("Processing data");
                                        await ProcessReceivedDataAsync(receivedData);
                                    }
                                    catch (IOException)
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    catch (SocketException ex)
                    {
                        _eventLog1.WriteEntry($"Connection failed: {ex.Message}. Retrying in 5 seconds...", EventLogEntryType.Warning);
                        await Task.Delay(5000, cancellationToken);
                    }
                }
            }
            catch (Exception ex)
            {
                _eventLog1.WriteEntry($"Fatal error in TCP client: {ex.Message}", EventLogEntryType.Error);
            }
        }

        /// <summary>
        /// Starts monitoring the configured directory for new files.
        /// Processes newly detected files asynchronously.
        /// </summary>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        private async Task StartFileSystemWatcherAsync(CancellationToken cancellationToken)
        {
            try
            {
                string dataDir = GetConfigValue("DATA_REPORT_DIRECTORY");
                FileSystemWatcher watcher = new FileSystemWatcher(dataDir)
                {
                    Filter = "*.*",
                    EnableRaisingEvents = true
                };

                watcher.Created += async (sender, e) => await HandleDataFileAsync(sender, e, cancellationToken);

                await Task.Delay(Timeout.Infinite, cancellationToken);
            }
            catch (Exception ex)
            {
                _eventLog1.WriteEntry($"Fatal error while using the file system watcher: {ex.Message}", EventLogEntryType.Error);
            }
        }

        /// <summary>
        /// Handles new files detected in the monitored directory.
        /// Implements a retry mechanism to handle file locks.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">File system event arguments containing file details.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        private async Task HandleDataFileAsync(object sender, FileSystemEventArgs e, CancellationToken cancellationToken)
        {
            const int maxRetries = 5;
            const int delayBetweenRetriesMs = 1000; // 1 second

            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    _eventLog1.WriteEntry($"Attempting to process file: {e.FullPath} (Attempt {attempt + 1}/{maxRetries})");

                    // Check if the file is still locked
                    if (IsFileLocked(e.FullPath))
                    {
                        _eventLog1.WriteEntry($"File is still being used by another process: {e.FullPath}", EventLogEntryType.Warning);
                        await Task.Delay(delayBetweenRetriesMs, cancellationToken);
                        continue;
                    }

                    string fileContent;

                    // Read the file if it's ready
                    using (var fileStream = new FileStream(e.FullPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    using (var reader = new StreamReader(fileStream))
                    {
                        fileContent = await reader.ReadToEndAsync();
                    }

                    // Process the file content
                    await ProcessReceivedDataAsync(fileContent);
                    _eventLog1.WriteEntry($"Successfully processed file: {e.FullPath}");
                    break; // Exit loop if successful
                }
                catch (IOException ioEx)
                {
                    _eventLog1.WriteEntry($"IO Error processing file {e.FullPath}: {ioEx.Message}. Retrying...", EventLogEntryType.Warning);
                    await Task.Delay(delayBetweenRetriesMs, cancellationToken);
                }
                catch (Exception ex)
                {
                    _eventLog1.WriteEntry($"Unexpected error processing file {e.FullPath}: {ex.Message}", EventLogEntryType.Error);
                    break; // Exit on unexpected errors
                }
            }
        }

        /// <summary>
        /// Checks if the specified file is currently locked by another process.
        /// </summary>
        /// <param name="filePath">The full path of the file to check.</param>
        /// <returns>True if the file is locked, otherwise false.</returns>
        private bool IsFileLocked(string filePath)
        {
            try
            {
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    // If successful, file is not locked
                    return false;
                }
            }
            catch (IOException)
            {
                // IOException indicates file is still in use
                return true;
            }
        }

        /// <summary>
        /// Processes the received data by parsing probe numbers, measurement dates, comments, and metal values.
        /// Validates the data and stores it in the appropriate SQL database.
        /// </summary>
        /// <param name="data">The raw data string to process.</param>
        private async Task ProcessReceivedDataAsync(string data)
        {
            try
            {
                _eventLog1.WriteEntry($"Received data: {data}");

                var germanCulture = new CultureInfo("de-DE");

                string orderNo = string.Empty;
                string comment = string.Empty;
                DateTime? measureDate = null;
                decimal au = 0, ag = 0, pt = 0, pd = 0, rh = 0;
                string targetDatabase = "TESTDB"; // Default database

                var lines = data.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var line in lines)
                {
                    if (line.StartsWith("%Probe:"))
                    {
                        string probeValue = line.Substring("%Probe:".Length).Trim();
                        _eventLog1.WriteEntry($"Probe number: {probeValue}");

                        if (probeValue.StartsWith("B-") || probeValue.StartsWith("B "))
                        {
                            targetDatabase = "TESTDB1";
                            orderNo = probeValue.Substring(2);
                        }
                        else if (probeValue.StartsWith("G-") || probeValue.StartsWith("G "))
                        {
                            targetDatabase = "TESTDB2";
                            orderNo = probeValue.Substring(2);
                        }
                        else
                        {
                            targetDatabase = "TESTDB3";
                            orderNo = probeValue;
                        }
                    }
                    else if (line.StartsWith("%Bemerkung:"))
                    {
                        comment = line.Substring("%Bemerkung:".Length).Trim();
                    }
                    else if (line.StartsWith("%Datum:"))
                    {
                        string dateStr = line.Substring("%Datum:".Length).Trim();
                        if (DateTime.TryParseExact(dateStr, "dd.MM.yyyy", germanCulture, DateTimeStyles.None, out DateTime parsedDate))
                        {
                            measureDate = parsedDate;
                        }
                        else
                        {
                            _eventLog1.WriteEntry($"Failed to parse measure date: {dateStr}", EventLogEntryType.Warning);
                        }
                    }
                    else if (line.Contains(";"))
                    {
                        var parts = line.Split(';');
                        if (parts.Length >= 3)
                        {
                            string element = parts[0].Trim();
                            string value = parts[2].Trim();

                            if (string.IsNullOrWhiteSpace(value) || value == "-")
                            {
                                value = "0,0";
                            }

                            switch (element)
                            {
                                case "Au":
                                    if (!decimal.TryParse(value, NumberStyles.Any, germanCulture, out au))
                                    {
                                        _eventLog1.WriteEntry($"Failed to parse Au value: {value}, defaulting to 0.0", EventLogEntryType.Warning);
                                        au = 0;
                                    }
                                    break;
                                case "Ag":
                                    if (!decimal.TryParse(value, NumberStyles.Any, germanCulture, out ag))
                                    {
                                        _eventLog1.WriteEntry($"Failed to parse Ag value: {value}, defaulting to 0.0", EventLogEntryType.Warning);
                                        ag = 0;
                                    }
                                    break;
                                case "Pt":
                                    if (!decimal.TryParse(value, NumberStyles.Any, germanCulture, out pt))
                                    {
                                        _eventLog1.WriteEntry($"Failed to parse Pt value: {value}, defaulting to 0.0", EventLogEntryType.Warning);
                                        pt = 0;
                                    }
                                    break;
                                case "Pd":
                                    if (!decimal.TryParse(value, NumberStyles.Any, germanCulture, out pd))
                                    {
                                        _eventLog1.WriteEntry($"Failed to parse Pd value: {value}, defaulting to 0.0", EventLogEntryType.Warning);
                                        pd = 0;
                                    }
                                    break;
                                case "Rh":
                                    if (!decimal.TryParse(value, NumberStyles.Any, germanCulture, out rh))
                                    {
                                        _eventLog1.WriteEntry($"Failed to parse Rh value: {value}, defaulting to 0.0", EventLogEntryType.Warning);
                                        rh = 0;
                                    }
                                    break;
                            }
                        }
                    }
                }

                // Validation: Ensure required fields are present
                if (string.IsNullOrWhiteSpace(orderNo))
                {
                    _eventLog1.WriteEntry("Missing probe number. Data will not be stored.", EventLogEntryType.Warning);
                    return;
                }

                if (!measureDate.HasValue)
                {
                    _eventLog1.WriteEntry("Missing measurement date. Data will not be stored.", EventLogEntryType.Warning);
                    return;
                }

                if (au == 0 && ag == 0 && pt == 0 && pd == 0 && rh == 0)
                {
                    _eventLog1.WriteEntry("No metal measurement values found. Data will not be stored.", EventLogEntryType.Warning);
                    return;
                }

                try
                {
                    _eventLog1.WriteEntry($"Storing data in MSSQL DB {targetDatabase}" +
                        $"\nOrder: {orderNo}\nMeasurement Date: {measureDate}\nAu: {au}\nAg: {ag}\nPt: {pt}\nPd: {pd}\nRh: {rh}");
                    await StoreDataInMsSqlAsync(targetDatabase, orderNo, comment, au, ag, pt, pd, rh, measureDate);
                }
                catch (Exception ex)
                {
                    _eventLog1.WriteEntry($"MS SQL storage failed: {ex.Message}", EventLogEntryType.Warning);
                }
            }
            catch (Exception ex)
            {
                _eventLog1.WriteEntry($"Error processing received data: {ex.Message}", EventLogEntryType.Error);
            }
        }

        /// <summary>
        /// Stores the processed measurement data into the specified SQL Server database.
        /// Updates the existing records with new measurement values.
        /// </summary>
        /// <param name="database">The target database name.</param>
        /// <param name="orderNo">The probe number (order number).</param>
        /// <param name="comment">Any comments associated with the measurement.</param>
        /// <param name="au">Gold measurement value.</param>
        /// <param name="ag">Silver measurement value.</param>
        /// <param name="pt">Platinum measurement value.</param>
        /// <param name="pd">Palladium measurement value.</param>
        /// <param name="rh">Rhodium measurement value.</param>
        /// <param name="measureDate">The date of the measurement.</param>
        private async Task StoreDataInMsSqlAsync(string database, string orderNo, string comment, decimal au, decimal ag, decimal pt, decimal pd, decimal rh, DateTime? measureDate)
        {
            string connectionString = GetDatabaseConnectionString(database);
            _eventLog1.WriteEntry($"DB Connection: {connectionString}");

            using (var connection = new SqlConnection(connectionString))
            {
                await connection.OpenAsync();

                using (var command = connection.CreateCommand())
                {
                    command.CommandText = @"
                        UPDATE tblScheidgut_Auftrag 
                        SET curRFAAu = @au,
                            curRFAAg = @ag,
                            curRFAPt = @pt,
                            curRFAPd = @pd,
                            curRFARh = @rh,
                            strRFAComment = @comment,
                            dtmRFAMeasureDate = @measureDate,
                            dtmRFACreatedAt = @createdAt
                        WHERE PostenNr = @orderNo";

                    command.Parameters.AddWithValue("@orderNo", orderNo);
                    command.Parameters.AddWithValue("@comment", (object)comment ?? DBNull.Value);
                    command.Parameters.AddWithValue("@au", au);
                    command.Parameters.AddWithValue("@ag", ag);
                    command.Parameters.AddWithValue("@pt", pt);
                    command.Parameters.AddWithValue("@pd", pd);
                    command.Parameters.AddWithValue("@rh", rh);
                    command.Parameters.AddWithValue("@measureDate", measureDate);
                    command.Parameters.AddWithValue("@createdAt", DateTime.UtcNow);

                    await command.ExecuteNonQueryAsync();
                }
            }

            _eventLog1.WriteEntry($"Successfully stored data for order {orderNo} in MS SQL database {database}", EventLogEntryType.Information);
        }

        private void eventLog1_EntryWritten(object sender, EntryWrittenEventArgs e)
        {
        }

        /// <summary>
        /// Retrieves the connection string for the specified database.
        /// Falls back to the default connection string if no specific one is found.
        /// </summary>
        /// <param name="database">The target database name.</param>
        /// <returns>The connection string for the database.</returns>
        private string GetDatabaseConnectionString(string database)
        {
            string configKey = $"MSSQL_{database.ToUpper()}_CONNECTION_STRING";
            string connectionString = GetConfigValue(configKey, required: false);

            if (string.IsNullOrEmpty(connectionString))
            {
                _eventLog1.WriteEntry($"No specific connection string found for database {database}, using default connection string", EventLogEntryType.Warning);
                connectionString = GetConfigValue("MSSQL_CONNECTION_STRING");
            }

            return connectionString;
        }

        /// <summary>
        /// Retrieves the configuration value from app settings or environment variables.
        /// Throws an exception if the required key is missing.
        /// </summary>
        /// <param name="key">The configuration key to retrieve.</param>
        /// <param name="required">Indicates if the key is mandatory.</param>
        /// <returns>The value of the configuration key.</returns>
        private string GetConfigValue(string key, bool required = true)
        {
            string value = ConfigurationManager.AppSettings[key];
            if (string.IsNullOrEmpty(value))
            {
                value = Environment.GetEnvironmentVariable(key);
            }
            if (string.IsNullOrEmpty(value) && required)
            {
                throw new ConfigurationErrorsException($"Configuration for {key} not found in app settings or environment variables.");
            }
            return value;
        }
    }
}
