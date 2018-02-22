# Prog
A PowerShell module for logging messages to the host, to PowerShell streams, or to a log file.

## Features

### Writing Log Messages
Messages are logged using the **Write-LogMessage** function.

### Configuration
The Logging module may be configured prior to writing any log messages, using the function **Set-LogConfiguration**.  The current configuration can be read using the function **Get-LogConfiguration**.  Function **Reset-LogConfiguration** will set the configuration back to its default settings.

**Set-LogConfiguration** may be used to set the following log properties:

1) **The log level:**  This determines whether a message will be logged or not.  
	
   Possible log levels, in order from lowest to highest, are: 
   * Off
   * Error
   * Warning
   * Information
   * Debug
   * Verbose 

   Only log messages at a level the same as, or lower than, the LogLevel will be logged.  For example, if the LogLevel is "Information" then only log messages at a level of Information, Warning or Error will be logged.  Messages at a level of Debug or Verbose will not be logged, as these log levels are higher than Information;

2) **The message destination:**  Messages may be written to the host or to PowerShell streams such as the Information stream or the Verbose stream.  In addition, if a log file name is set in the configuration, the messages will be written to a log file;

3) **The host text color:**  Messages written to the host, as opposed to PowerShell streams, may be written in any PowerShell console color.  Different colors may be specified for different message types, such as Error, Warning or Information;

4) **The message format:**  In addition to the specified message, the text written to the log may include additional fields that are automatically populated, such as a timestamp or the name of the function writing to the log.  A simple template can be defined to specify the format of the logged text, including the fields to be displayed and any field separators.

Some log properties can be overridden when writing a single log message.  The changes apply only to that one message; subsequent messages will return to using the settings specified via Set-LogConfiguration.