# Variables that represent the internal state of the Pslogg module.

$_logLevels = @{
                    Off = 0
                    Error = 1
                    Warning = 2
                    Information = 3
                    Debug = 4
                    Verbose = 5
                }

$_defaultHostTextColor = @{
                                Error = 'Red'
                                Warning = 'Yellow'
                                Information = 'White'
                                Debug = 'White'
                                Verbose = 'White'
                            }

$_defaultCategoryInfo = @{
                        Progress = @{ IsDefault = $True }
                        Success = @{ Color = 'Green' }
                        Failure = @{ Color = 'Red' }
                        PartialFailure = @{ Color = 'Yellow' }
                    }

$_defaultLogConfiguration = @{   
                                LogLevel = 'INFORMATION'
								MessageFormat = '{Timestamp:yyyy-MM-dd HH:mm:ss.fff} | {CallerName} | {Category} | {MessageLevel} | {Message}'
                                WriteToHost = $True
                                HostTextColor = $_defaultHostTextColor
                                LogFile = @{
                                                WriteFromScript = $True
                                                WriteFromHost = $False
                                                Name = 'Results.log'
                                                IncludeDateInFileName = $True
                                                Overwrite = $True
                                                FullPathReadOnly = ''
                                            }
                                CategoryInfo = $_defaultCategoryInfo
                            }

$_defaultTimestampFormat = 'yyyy-MM-dd HH:mm:ss.fff'	
						
$_logConfiguration = @{}
$_messageFormatInfo = @{}

$_logFileOverwritten = $False