Option Explicit

''' Enum AdoDb Connnection, Command & Recordset StateEnum
Const adStateClosed = 0 'Object is closed
Const adStateOpen = 1 'Object is open
Const adStateConnecting = 2 'Object is connecting
Const adStateExecuting = 4 'Object is executing ... n/a for Connection
Const adStateFetching = 8 'Object is fetching ... n/a for Connection
''' End Enum

''' Enum AdoDb SchemaEnum Constants
Const adSchemaColumns = 4 'Request column information
Const adSchemaProviderTypes = 22 'Request provider type information
Const adSchemaTables = 20 'Request table information
''' End Enum

''' Enum I/O Mode Constants (fso.OpenTextFile)
Const ForReading = 1 'Opens a file for reading only
Const ForWriting = 2 'Opens a file for writing. If the file already exists, the contents are overwritten.
Const ForAppending = 8 'Opens a file and starts writing at the end (appends). Contents are not overwritten.
''' End Enum

''' Enum Tristate Constants
Const TristateFalse = 0 'Opens the file as ASCII
Const TristateTrue = - 1 'Opens the file as Unicode
Const TristateUseDefault = - 2 'Use default system setting
''' End Enum