# RFAConnectorService Documentation

## Overview
The RFA Connector Service is a Windows service that retrieves measurement data from [FISCHERSCOPE® X-RAY XRF devices](https://www.helmut-fischer.com/products/xrf-instruments/) and stores it in a database. It connects to a TCP server (Fischer WinFTM with TCP/IP Server enabled) or monitors a directory for incoming data files. It processes measurement data and stores it in a Microsoft SQL Server database. The service supports two operational modes:

1. **TCP Client Mode**: Connects to a remote server to receive data.
2. **File System Watcher Mode**: Monitors a directory for new data files.

>[!IMPORTANT]
>Reading the measurement data requires a specific export format from WinFTM. If your export template deviates from this schema, the parser in the code must be adjusted accordingly.

The service can distinguish between different clients based on the template prefixes and store the measurement values in the correct target database.

The service has its own setup program. The service is automatically started after installation. 

## Configuration
The service relies on several configuration values that can be set in the application settings or environment variables.

| Configuration Key                   | Description                                                     | Required |
|-------------------------------------|-----------------------------------------------------------------|----------|
| `ENABLE_TCP_CONNECTION`            | Enables TCP connection if set to `true`, otherwise uses file watcher. | Yes      |
| `RFA_TCP_HOST`                     | The hostname of the TCP server.                                  | Yes (if TCP enabled) |
| `RFA_TCP_PORT`                     | The port of the TCP server.                                      | Yes (if TCP enabled) |
| `DATA_REPORT_DIRECTORY`            | Directory path to monitor for new data files.                    | Yes (if TCP disabled) |
| `MSSQL_CONNECTION_STRING`          | Default connection string for the SQL database.                  | Yes      |
| `MSSQL_DB1_CONNECTION_STRING`  | Connection string for the 'DB1' database.                     | No       |
| `MSSQL_DB2_CONNECTION_STRING`      | Connection string for the 'DB2' database.                        | No       |
| `MSSQL_DB3_CONNECTION_STRING`  | Connection string for the 'DB3' database.                    | No       |

## Service Lifecycle

### OnStart
- Initializes the event log.
- Determines the operation mode (TCP client or file watcher).
- Starts the appropriate data receiver.

### OnStop
- Cancels any running tasks.
- Closes the TCP connection if active.
- Updates the service status to stopped.

## Core Components

### Enums and Structs
- **ServiceState**: Defines the different states of the service.
- **ServiceStatus**: Represents the status structure used to communicate with the Windows Service Control Manager.

### Methods

#### StartTcpClientAsync
- Connects to the configured TCP server.
- Reads incoming data and processes it using `ProcessReceivedDataAsync`.
- Reconnects automatically on connection failure.

#### StartFileSystemWatcherAsync
- Monitors the configured directory for new files.
- Processes newly detected files using `HandleDataFileAsync`.

#### HandleDataFileAsync
- Attempts to read the file with retries if the file is locked by another process.
- Calls `ProcessReceivedDataAsync` to handle the file content.

#### IsFileLocked
- Checks if a file is currently being used by another process.

#### ProcessReceivedDataAsync
- Parses received data (from TCP or file).
- Extracts probe number, measurement date, comments, and metal values (Au, Ag, Pt, Pd, Rh).
- Validates that essential data (probe number, measurement date, and at least one metal value) is present.
- Calls `StoreDataInMsSqlAsync` to store the processed data.

#### StoreDataInMsSqlAsync
- Connects to the appropriate SQL Server database.
- Updates measurement data in the `tblScheidgut_Auftrag` table.

#### GetDatabaseConnectionString
- Retrieves the connection string for the specified database.
- Falls back to the default connection string if a specific one is not found.

#### GetConfigValue
- Retrieves configuration values from app settings or environment variables.
- Throws an exception if a required value is missing.

## Logging
The service logs important events and errors to the Windows Event Log under `RFAConnectorLog` with the source `RFAConnectorSource`. Logged events include:

- Service start and stop events.
- TCP connection status and errors.
- File processing attempts and errors.
- Data validation failures.
- Database connection and storage status.

## Error Handling
- **TCP Connection Errors**: Retries connection every 5 seconds if it fails.
- **File Access Errors**: Retries reading files up to 5 times if they are locked.
- **Data Validation Errors**: Logs warnings if essential data (probe number, measurement date, or metal values) is missing and skips storage.
- **Database Errors**: Logs warnings if database operations fail.

## Database Schema
The processed data is stored in the `tblScheidgut_Auftrag` table with the following columns:

| Column Name        | Data Type | Description                                  |
|--------------------|-----------|----------------------------------------------|
| `PostenNr`         | String    | The probe number (order number).             |
| `curRFAAu`         | Decimal   | Gold measurement value.                      |
| `curRFAAg`         | Decimal   | Silver measurement value.                    |
| `curRFAPt`         | Decimal   | Platinum measurement value.                  |
| `curRFAPd`         | Decimal   | Palladium measurement value.                 |
| `curRFARh`         | Decimal   | Rhodium measurement value.                   |
| `strRFAComment`    | String    | Comment associated with the measurement.     |
| `dtmRFAMeasureDate`| DateTime  | Date of the measurement.                     |
| `dtmRFACreatedAt`  | DateTime  | Timestamp when the data was processed.       |

## Example Data Format

```
%Probe: B-12345
%Bemerkung: Sample processed
%Datum: 12.02.2025
Au;some;12,5
Ag;some;8,3
Pt;some;-
Pd;some;0,0
Rh;some;1,2
```

### Parsed Output
- **Probe Number**: 12345
- **Comment**: Sample processed
- **Measurement Date**: 12.02.2025
- **Au**: 12.5
- **Ag**: 8.3
- **Pt**: 0.0
- **Pd**: 0.0
- **Rh**: 1.2

## Troubleshooting
1. **File Locked Errors**:
   - Ensure the process writing files releases them properly.
   - The service will retry reading locked files up to 5 times.

2. **Database Connection Issues**:
   - Verify connection strings in configuration.
   - Ensure the database server is reachable.

3. **Missing Data Warnings**:
   - Ensure incoming data includes probe number, measurement date, and at least one metal measurement.

4. **TCP Connection Failures**:
   - Verify the TCP server host and port configurations.
   - Ensure network connectivity.

---


