
Imports System.Runtime.InteropServices
Imports System.Text

Namespace Cryptlex
    Public NotInheritable Class LexActivator
        Private Sub New()
        End Sub
        Private Const DLL_FILE_NAME As String = "LexActivator.dll"

        '
        '           In order to use "Any CPU" configuration, rename 64 bit LexActivator.dll to LexActivator64.dll and add "LA_ANY_CPU"
        '	        conditional compilation symbol.
        '        
        '           #Const LA_ANY_CPU = 1
#If LA_ANY_CPU Then
        Private Const DLL_FILE_NAME_X64 As String = "LexActivator64.dll"
#End If
        Public Enum PermissionFlags As UInteger
            LA_USER = 1
            LA_SYSTEM = 2
        End Enum

        Public Enum TrialType As UInteger
            LA_V_TRIAL = 1
            LA_UV_TRIAL = 2
        End Enum

        '
        '            FUNCTION: SetProductFile()
        '
        '            PURPOSE: Sets the path of the Product.dat file. This should be
        '            used if your application and Product.dat file are in different
        '            folders or you have renamed the Product.dat file.
        '
        '            If this function is used, it must be called on every start of
        '            your program before any other functions are called.
        '
        '            PARAMETERS:
        '            * filePath - path of the product file (Product.dat)
        '
        '            RETURN CODES: LA_OK, LA_E_FPATH, LA_E_PFILE
        '
        '            NOTE: If this function fails to set the path of product file, none of the
        '            other functions will work.
        '        


        Public Shared Function SetProductFile(filePath As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.SetProductFile_x64(filePath), Native.SetProductFile(filePath))
#Else
            Return Native.SetProductFile(filePath)
#End If

        End Function

        '
        '            FUNCTION: SetVersionGUID()
        '
        '            PURPOSE: Sets the version GUID of your application. 
        '
        '            This function must be called on every start of your program before
        '            any other functions are called, with the exception of SetProductFile()
        '            function.
        '
        '            PARAMETERS:
        '            * versionGUID - the unique version GUID of your application as mentioned
        '              on the product version page of your application in the dashboard.
        '
        '            * flags - depending upon whether your application requires admin/root 
        '              permissions to run or not, this parameter can have one of the following
        '              values: LA_SYSTEM, LA_USER
        '
        '            RETURN CODES: LA_OK, LA_E_WMIC, LA_E_PFILE, LA_E_GUID, LA_E_PERMISSION
        '
        '            NOTE: If this function fails to set the version GUID, none of the other
        '            functions will work.
        '        


        Public Shared Function SetVersionGUID(versionGUID As String, flags As PermissionFlags) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.SetVersionGUID_x64(versionGUID, flags), Native.SetVersionGUID(versionGUID, flags))
#Else
            Return Native.SetVersionGUID(versionGUID, flags)
#End If

        End Function

        '
        '            FUNCTION: SetProductKey()
        '
        '            PURPOSE: Sets the product key required to activate the application.
        '
        '            PARAMETERS:
        '            * productKey - a valid product key generated for the application.
        '
        '            RETURN CODES: LA_OK, LA_E_GUID, LA_E_PKEY
        '        


        Public Shared Function SetProductKey(productKey As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.SetProductKey_x64(productKey), Native.SetProductKey(productKey))
#Else
            Return Native.SetProductKey(productKey)
#End If

        End Function

        '
        '            FUNCTION: SetExtraActivationData()
        '
        '            PURPOSE: Sets the extra data which you may want to fetch from the user
        '            at the time of activation. 
        '
        '            The extra data appears along with activation details of the product key
        '            in dashboard.
        '
        '            PARAMETERS:
        '            * extraData - string of maximum length 256 characters with utf-8 encoding.
        '
        '            RETURN CODES: LA_OK, LA_E_EDATA_LEN, LA_E_GUID
        '
        '            NOTE: If the length of the string is more than 256, it is truncated to the
        '            allowed size.
        '        


        Public Shared Function SetExtraActivationData(extraData As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.SetExtraActivationData_x64(extraData), Native.SetExtraActivationData(extraData))
#Else
            Return Native.SetExtraActivationData(extraData)
#End If

        End Function

        '
        '            FUNCTION: ActivateProduct()
        '
        '            PURPOSE: Activates your application by contacting the Cryptlex servers. It 
        '            validates the key and returns with encrypted and digitally signed response
        '            which it stores and uses to activate your application.
        '
        '            This function should be executed at the time of registration, ideally on
        '            a button click. 
        '
        '            RETURN CODES: LA_OK, LA_EXPIRED, LA_REVOKED, LA_FAIL, LA_E_GUID, LA_E_PKEY,
        '            LA_E_INET, LA_E_VM, LA_E_TIME, LA_E_ACT_LIMIT
        '        


        Public Shared Function ActivateProduct() As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.ActivateProduct_x64(), Native.ActivateProduct())
#Else
            Return Native.ActivateProduct()
#End If

        End Function

        '
        '            FUNCTION: DeactivateProduct()
        '
        '            PURPOSE: Deactivates the application and frees up the correponding activation
        '            slot by contacting the Cryptlex servers.
        '
        '            This function should be executed at the time of deregistration, ideally on
        '            a button click. 
        '
        '            RETURN CODES: LA_OK, LA_EXPIRED, LA_REVOKED, LA_FAIL, LA_E_GUID, LA_E_PKEY,
        '            LA_E_INET, LA_E_DEACT_LIMIT
        '        


        Public Shared Function DeactivateProduct() As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.DeactivateProduct_x64(), Native.DeactivateProduct())
#Else
            Return Native.DeactivateProduct()
#End If
        End Function

        '
        '            FUNCTION: ActivateProductOffline()
        '
        '            PURPOSE: Activates your application using the offline activation response
        '            file.
        '
        '            RETURN CODES: LA_OK, LA_EXPIRED, LA_FAIL, LA_E_GUID, LA_E_PKEY, LA_E_OFILE
        '            LA_E_VM, LA_E_TIME
        '        


        Public Shared Function ActivateProductOffline(filePath As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.ActivateProductOffline_x64(filePath), Native.ActivateProductOffline(filePath))
#Else
            Return Native.ActivateProductOffline(filePath)
#End If
        End Function

        '
        '            FUNCTION: GenerateOfflineActivationRequest()
        '
        '            PURPOSE: Generates the offline activation request needed for generating
        '            offline activation response in the dashboard.
        '
        '            PARAMETERS:
        '            * filePath - path of the file, needed to be created, for the offline request.
        '
        '            RETURN CODES: LA_OK, LA_FAIL, LA_E_GUID, LA_E_PKEY
        '        


        Public Shared Function GenerateOfflineActivationRequest(filePath As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GenerateOfflineActivationRequest_x64(filePath), Native.GenerateOfflineActivationRequest(filePath))
#Else
            Return Native.GenerateOfflineActivationRequest(filePath)
#End If
        End Function

        '
        '            FUNCTION: GenerateOfflineDeactivationRequest()
        '
        '            PURPOSE: Generates the offline deactivation request needed for deactivation of 
        '            the product key in the dashboard and deactivates the application.
        '
        '            A valid offline deactivation file confirms that the application has been successfully
        '            deactivated on the user's machine.
        '
        '            PARAMETERS:
        '            * filePath - path of the file, needed to be created, for the offline request.
        '
        '            RETURN CODES: LA_OK, LA_FAIL, LA_E_GUID, LA_E_PKEY
        '        


        Public Shared Function GenerateOfflineDeactivationRequest(filePath As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GenerateOfflineDeactivationRequest_x64(filePath), Native.GenerateOfflineDeactivationRequest(filePath))
#Else
            Return Native.GenerateOfflineDeactivationRequest(filePath)
#End If
        End Function

        '
        '        FUNCTION: IsProductGenuine()
        '
        '        PURPOSE: It verifies whether your app is genuinely activated or not. The verfication is
        '        done locally by verifying the cryptographic digital signature fetched at the time of
        '        activation.
        '
        '        After verifying locally, it schedules a server check in a separate thread on due dates.
        '        The default interval for server check is 100 days and this can be changed if required.
        '
        '        In case server validation fails due to network error, it retries every 15 minutes. If it
        '        continues to fail for fixed number of days (grace period), the function returns LA_GP_OVER
        '        instead of LA_OK. The default length of grace period is 30 days and this can be changed if 
        '        required.
        '
        '        This function must be called on every start of your program to verify the activation
        '        of your app.
        '
        '        RETURN CODES: LA_OK, LA_EXPIRED, LA_REVOKED, LA_GP_OVER, LA_FAIL, LA_E_GUID, LA_E_PKEY
        '
        '        NOTE: If application was activated offline using ActivateProductOffline() function, you
        '        may want to set grace period to 0 to ignore grace period.
        '    


        Public Shared Function IsProductGenuine() As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.IsProductGenuine_x64(), Native.IsProductGenuine())
#Else
            Return Native.IsProductGenuine()
#End If
        End Function

        '
        '            FUNCTION: IsProductActivated()
        '
        '            PURPOSE: It verifies whether your app is genuinely activated or not. The verfication is
        '            done locally by verifying the cryptographic digital signature fetched at the time of
        '            activation.
        '
        '            This is just an auxillary function which you may use in some specific cases.
        '
        '            RETURN CODES: LA_OK, LA_EXPIRED, LA_REVOKED, LA_GP_OVER, LA_FAIL, LA_E_GUID, LA_E_PKEY
        '        


        Public Shared Function IsProductActivated() As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.IsProductActivated_x64(), Native.IsProductActivated())
#Else
            Return Native.IsProductActivated()
#End If
        End Function

        '
        '            FUNCTION: GetExtraActivationData()
        '
        '            PURPOSE: Gets the value of the extra activation data.
        '
        '            PARAMETERS:
        '            * extraData - pointer to a buffer that receives the value of the string
        '            * length - size of the buffer pointed to by the fieldValue parameter
        '
        '            RETURN CODES: LA_OK, LA_E_GUID, LA_E_BUFFER_SIZE
        '        


        Public Shared Function GetExtraActivationData(extraData As StringBuilder, length As Integer) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GetExtraActivationData_x64(extraData, length), Native.GetExtraActivationData(extraData, length))
#Else
            Return Native.GetExtraActivationData(extraData, length)
#End If
        End Function

        '
        '            FUNCTION: GetCustomLicenseField()
        '
        '            PURPOSE: Gets the value of the custom field associated with the product key.
        '
        '            PARAMETERS:
        '            * fieldId - id of the custom field whose value you want to get
        '            * fieldValue - pointer to a buffer that receives the value of the string
        '            * length - size of the buffer pointed to by the fieldValue parameter
        '
        '            RETURN CODES: LA_OK, LA_E_CUSTOM_FIELD_ID, LA_E_GUID, LA_E_BUFFER_SIZE
        '        


        Public Shared Function GetCustomLicenseField(fieldId As String, fieldValue As StringBuilder, length As Integer) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GetCustomLicenseField_x64(fieldId, fieldValue, length), Native.GetCustomLicenseField(fieldId, fieldValue, length))
#Else
            Return Native.GetCustomLicenseField(fieldId, fieldValue, length)
#End If
        End Function

        '
        '            FUNCTION: GetProductKey()
        '
        '            PURPOSE: Gets the stored product key which was used for activation.
        '
        '            PARAMETERS:
        '            * productKey - pointer to a buffer that receives the value of the string
        '            * length - size of the buffer pointed to by the productKey parameter
        '
        '            RETURN CODES: LA_OK, LA_E_PKEY, LA_E_GUID, LA_E_BUFFER_SIZE
        '        


        Public Shared Function GetProductKey(productKey As StringBuilder, length As Integer) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GetProductKey_x64(productKey, length), Native.GetProductKey(productKey, length))
#Else
            Return Native.GetProductKey(productKey, length)
#End If
        End Function

        '
        '            FUNCTION: GetDaysLeftToExpiration()
        '
        '            PURPOSE: Gets the number of remaining days after which the license expires.
        '
        '            PARAMETERS:
        '            * daysLeft - pointer to the integer that receives the value
        '            
        '            RETURN CODES: LA_OK, LA_FAIL, LA_E_GUID
        '


        Public Shared Function GetDaysLeftToExpiration(ByRef daysLeft As UInteger) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GetDaysLeftToExpiration_x64(daysLeft), Native.GetDaysLeftToExpiration(daysLeft))
#Else
            Return Native.GetDaysLeftToExpiration(daysLeft)
#End If
        End Function

        '
        '            FUNCTION: GetProductKeyExpiryDate()
        '
        '            PURPOSE: Gets the timestamp of the expiry date.
        '
        '            PARAMETERS:
        '            * expiryDate - pointer to the integer that receives the value
        '            
        '            RETURN CODES: LA_OK, LA_FAIL, LA_E_GUID
        '


        Public Shared Function GetProductKeyExpiryDate(ByRef expiryDate As UInteger) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GetProductKeyExpiryDate_x64(expiryDate), Native.GetProductKeyExpiryDate(expiryDate))
#Else
            Return Native.GetProductKeyExpiryDate(expiryDate)
#End If
        End Function


        '
        '            FUNCTION: SetTrialKey()
        '
        '            PURPOSE: Sets the trial key required to activate the verified trial.
        '
        '            PARAMETERS:
        '            * trialKey - trial key corresponding to the product version
        '
        '            RETURN CODES: LA_OK, LA_E_GUID, LA_E_TKEY
        '        


        Public Shared Function SetTrialKey(trialKey As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.SetTrialKey_x64(trialKey), Native.SetTrialKey(trialKey))
#Else
            Return Native.SetTrialKey(trialKey)
#End If
        End Function

        '
        '            FUNCTION: ActivateTrial()
        '
        '            PURPOSE: Starts the verified trial in your application by contacting the 
        '            Cryptlex servers. 
        '
        '            This function should be executed when your application starts first time on
        '            the user's computer, ideally on a button click. 
        '
        '            RETURN CODES: LA_OK, LA_T_EXPIRED, LA_FAIL, LA_E_GUID, LA_E_TKEY, LA_E_INET,
        '            LA_E_VM, LA_E_TIME
        '        


        Public Shared Function ActivateTrial() As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.ActivateTrial_x64(), Native.ActivateTrial())
#Else
            Return Native.ActivateTrial()
#End If
        End Function

        '
        '            FUNCTION: IsTrialGenuine()
        '
        '            PURPOSE: It verifies whether trial has started and is genuine or not. The
        '            verfication is done locally by verifying the cryptographic digital signature
        '            fetched at the time of trial activation.
        '
        '            This function must be called on every start of your program during the trial period.
        '
        '            RETURN CODES: LA_OK, LA_T_EXPIRED, LA_FAIL, LA_E_GUID, LA_E_TKEY
        '
        '            NOTE: The function is only meant for verified trials.
        '        


        Public Shared Function IsTrialGenuine() As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.IsTrialGenuine_x64(), Native.IsTrialGenuine())
#Else
            Return Native.IsTrialGenuine()
#End If
        End Function

        '
        '            FUNCTION: ExtendTrial()
        '
        '            PURPOSE: Extends the trial using the trial extension key generated in the dashboard
        '            for the product version.
        '
        '            PARAMETERS:
        '            * trialExtensionKey - trial extension key generated for the product version
        '
        '            RETURN CODES: LA_OK, LA_TEXT_EXPIRED, LA_FAIL, LA_E_GUID, LA_E_TEXT_KEY, LA_E_TKEY,
        '            LA_E_INET, LA_E_VM, LA_E_TIME
        '
        '            NOTE: The function is only meant for verified trials.
        '        


        Public Shared Function ExtendTrial(trialExtensionKey As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.ExtendTrial_x64(trialExtensionKey), Native.ExtendTrial(trialExtensionKey))
#Else
            Return Native.ExtendTrial(trialExtensionKey)
#End If
        End Function

        '
        '            FUNCTION: InitializeTrial()
        '
        '            PURPOSE: Starts the unverified trial if trial has not started yet and if
        '            trial has already started, it verifies the validity of trial.
        '
        '            This function must be called on every start of your program during the trial period.
        '
        '            PARAMETERS:
        '            * trialLength - trial length as set in the dashboard for the product version
        '
        '            RETURN CODES: LA_OK, LA_T_EXPIRED, LA_FAIL, LA_E_GUID, LA_E_TRIAL_LEN
        '
        '            NOTE: The function is only meant for unverified trials.
        '        


        Public Shared Function InitializeTrial(trialLength As UInteger) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.InitializeTrial_x64(trialLength), Native.InitializeTrial(trialLength))
#Else
            Return Native.InitializeTrial(trialLength)
#End If
        End Function

        '
        '            FUNCTION: GetTrialDaysLeft()
        '
        '            PURPOSE: Gets the number of remaining trial days.
        '
        '            If the trial has expired or has been tampered, daysLeft is set to 0 days.
        '
        '            PARAMETERS:
        '            * daysLeft - pointer to the integer that receives the value
        '            * trialType - depending upon whether your application uses verified trial or not,
        '              this parameter can have one of the following values: LA_V_TRIAL, LA_UV_TRIAL
        '
        '            RETURN CODES: LA_OK, LA_FAIL, LA_E_GUID
        '
        '            NOTE: The trial must be started by calling ActivateTrial() or  InitializeTrial() at least
        '            once in the past before calling this function.
        '        


        Public Shared Function GetTrialDaysLeft(ByRef daysLeft As UInteger, trialType As TrialType) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GetTrialDaysLeft_x64(daysLeft, trialType), Native.GetTrialDaysLeft(daysLeft, trialType))
#Else
            Return Native.GetTrialDaysLeft(daysLeft, trialType)
#End If
        End Function

        '
        '            FUNCTION: SetDayIntervalForServerCheck()
        '
        '            PURPOSE: Sets the interval for server checks done by IsProductGenuine() function.
        '
        '            To disable sever check pass 0 as the day interval.
        '
        '            PARAMETERS:
        '            * dayInterval - length of the interval in days
        '
        '            RETURN CODES: LA_OK, LA_E_GUID
        '        


        Public Shared Function SetDayIntervalForServerCheck(dayInterval As UInteger) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.SetDayIntervalForServerCheck_x64(dayInterval), Native.SetDayIntervalForServerCheck(dayInterval))
#Else
            Return Native.SetDayIntervalForServerCheck(dayInterval)
#End If
        End Function

        '
        '            FUNCTION: SetGracePeriodForNetworkError()
        '
        '            PURPOSE: Sets the grace period for failed re-validation requests sent
        '            by IsProductGenuine() function, caused due to network errors.
        '    
        '            It determines how long in days, should IsProductGenuine() function retry
        '            contacting CryptLex Servers, before returning LA_GP_OVER instead of LA_OK.
        '
        '            To ignore grace period pass 0 as the grace period. This may be useful in
        '            case of offline activations.
        '
        '            PARAMETERS:
        '            * gracePeriod - length of the grace period in days
        '
        '            RETURN CODES: LA_OK, LA_E_GUID
        '        


        Public Shared Function SetGracePeriodForNetworkError(gracePeriod As UInteger) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.SetGracePeriodForNetworkError_x64(gracePeriod), Native.SetGracePeriodForNetworkError(gracePeriod))
#Else
            Return Native.SetGracePeriodForNetworkError(gracePeriod)
#End If
        End Function

        '
        '            FUNCTION: SetNetworkProxy()
        '
        '            PURPOSE: Sets the network proxy to be used when contacting CryptLex servers.
        '
        '            The proxy format should be: [protocol://][username:password@]machine[:port]
        '
        '            Following are some examples of the valid proxy strings:
        '                - http://127.0.0.1:8000/
        '                - http://user:pass@127.0.0.1:8000/
        '                - socks5://127.0.0.1:8000/
        '
        '            PARAMETERS:
        '            * proxy - proxy string having correct proxy format
        '
        '            RETURN CODES: LA_OK, LA_E_NET_PROXY, LA_E_GUID
        '
        '            NOTE: Proxy settings of the computer are automatically detected. So, in most of the 
        '            cases you don't need to care whether your user is behind a proxy server or not.
        '        


        Public Shared Function SetNetworkProxy(proxy As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.SetNetworkProxy_x64(proxy), Native.SetNetworkProxy(proxy))
#Else
            Return Native.SetNetworkProxy(proxy)
#End If
        End Function


        '
        '            FUNCTION: SetUserLock()
        '
        '            PURPOSE: Enables the user locked licensing.
        ' 
        '            It adds an additional user lock to the product key. Activations by different users in
        '            the same OS are treated as separate activations.
        '
        '            PARAMETERS:
        '            * userLock - boolean value to enable or disable the lock
        '
        '            RETURN CODES: LA_OK, LA_E_GUID
        '
        '            NOTE: User lock is disabled by default. You should enable it in case your application
        '            is used through remote desktop services where multiple users access individual sessions
        '            on a single machine instance at the same time.
        '        


        Public Shared Function SetUserLock(userLock As Boolean) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.SetUserLock_x64(userLock), Native.SetUserLock(userLock))
#Else
            Return Native.SetUserLock(userLock)
#End If
        End Function


        '
        '            FUNCTION: SetCryptlexHost()
        '
        '            PURPOSE: In case you are running Cryptlex on a private web server, you can set the
        '            host for your private server.
        '
        '            PARAMETERS:
        '            * host - the address of the private web server running Cryptlex
        '
        '            RETURN CODES: LA_OK, LA_E_HOST_URL, LA_E_GUID
        '
        '            NOTE: This function should never be used unless you have opted for a private Cryptlex
        '            Server.
        '        


        Public Shared Function SetCryptlexHost(host As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.SetCryptlexHost_x64(host), Native.SetCryptlexHost(host))
#Else
            Return Native.SetCryptlexHost(host)
#End If
        End Function

        '** Return Codes **


        Public Const LA_OK As Integer = &H0

        Public Const LA_FAIL As Integer = &H1

        '
        '            CODE: LA_EXPIRED
        '
        '            MESSAGE: The product key has expired or system time has been tampered
        '            with. Ensure your date and time settings are correct.
        '        


        Public Const LA_EXPIRED As Integer = &H2

        '
        '            CODE: LA_REVOKED
        '
        '            MESSAGE: The product key has been revoked.
        '        


        Public Const LA_REVOKED As Integer = &H3

        '
        '            CODE: LA_GP_OVER
        '
        '            MESSAGE: The grace period is over.
        '        


        Public Const LA_GP_OVER As Integer = &H4

        '
        '            CODE: LA_E_INET
        '
        '            MESSAGE: Failed to connect to the server due to network error.
        '        


        Public Const LA_E_INET As Integer = &H5

        '
        '            CODE: LA_E_PKEY
        '
        '            MESSAGE: Invalid product key.
        '        


        Public Const LA_E_PKEY As Integer = &H6

        '
        '            CODE: LA_E_PFILE
        '
        '            MESSAGE: Invalid or corrupted product file.
        '        


        Public Const LA_E_PFILE As Integer = &H7

        '
        '            CODE: LA_E_FPATH
        '
        '            MESSAGE: Invalid product file path.
        '        


        Public Const LA_E_FPATH As Integer = &H8

        '
        '            CODE: LA_E_GUID
        '
        '            MESSAGE: The version GUID doesn't match that of the product file.
        '        


        Public Const LA_E_GUID As Integer = &H9

        '
        '            CODE: LA_E_OFILE
        '
        '            MESSAGE: Invalid offline activation response file.
        '        


        Public Const LA_E_OFILE As Integer = &HA

        '
        '            CODE: LA_E_PERMISSION
        '
        '            MESSAGE: Insufficent system permissions. Occurs when LA_SYSTEM flag is used
        '            but application is not run with admin privileges.
        '        


        Public Const LA_E_PERMISSION As Integer = &HB

        '
        '            CODE: LA_E_EDATA_LEN
        '
        '            MESSAGE: Extra activation data length is more than 256 characters.
        '        



        Public Const LA_E_EDATA_LEN As Integer = &HC

        '
        '            CODE: LA_E_TKEY
        '
        '            MESSAGE: The trial key doesn't match that of the product file.
        '        


        Public Const LA_E_TKEY As Integer = &HD

        '
        '            CODE: LA_E_TIME
        '
        '            MESSAGE: The system time has been tampered with. Ensure your date
        '            and time settings are correct.
        '        


        Public Const LA_E_TIME As Integer = &HE

        '
        '            CODE: LA_E_VM
        '
        '            MESSAGE: Application is being run inside a virtual machine / hypervisor,
        '            and activation has been disallowed in the VM.
        '            but
        '        


        Public Const LA_E_VM As Integer = &HF

        '
        '            CODE: LA_E_WMIC
        '
        '            MESSAGE: Fingerprint couldn't be generated because Windows Management 
        '            Instrumentation (WMI) service has been disabled. This error is specific
        '            to Windows only.
        '        


        Public Const LA_E_WMIC As Integer = &H10

        '
        '            CODE: LA_E_TEXT_KEY
        '
        '            MESSAGE: Invalid trial extension key.
        '        


        Public Const LA_E_TEXT_KEY As Integer = &H11

        '
        '            CODE: LA_E_TRIAL_LEN
        '
        '            MESSAGE: The trial length doesn't match that of the product file.
        '        


        Public Const LA_E_TRIAL_LEN As Integer = &H12

        '
        '            CODE: LA_T_EXPIRED
        '
        '            MESSAGE: The trial has expired or system time has been tampered
        '            with. Ensure your date and time settings are correct.
        '        


        Public Const LA_T_EXPIRED As Integer = &H13

        '
        '            CODE: LA_TEXT_EXPIRED
        '
        '            MESSAGE: The trial extension key being used has already expired or system
        '            time has been tampered with. Ensure your date and time settings are correct.
        '        


        Public Const LA_TEXT_EXPIRED As Integer = &H14

        '
        '            CODE: LA_E_BUFFER_SIZE
        '
        '            MESSAGE: The buffer size was smaller than required.
        '        


        Public Const LA_E_BUFFER_SIZE As Integer = &H15

        '
        '            CODE: LA_E_CUSTOM_FIELD_ID
        '
        '            MESSAGE: Invalid custom field id.
        '        


        Public Const LA_E_CUSTOM_FIELD_ID As Integer = &H16

        '
        '            CODE: LA_E_NET_PROXY
        '
        '            MESSAGE: Invalid network proxy.
        '        


        Public Const LA_E_NET_PROXY As Integer = &H17

        '
        '            CODE: LA_E_HOST_URL
        '
        '            MESSAGE: Invalid Cryptlex host url.
        '        


        Public Const LA_E_HOST_URL As Integer = &H18

        '
        '            CODE: LA_E_DEACT_LIMIT
        '
        '            MESSAGE: Deactivation limit for key has reached.
        '        


        Public Const LA_E_DEACT_LIMIT As Integer = &H19

        '
        '            CODE: LA_E_ACT_LIMIT
        '
        '            MESSAGE: Activation limit for key has reached.
        '        


        Public Const LA_E_ACT_LIMIT As Integer = &H1A


        Private NotInheritable Class Native
            Private Sub New()
            End Sub
            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function SetProductFile(filePath As String) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function SetVersionGUID(versionGUID As String, flags As PermissionFlags) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function SetProductKey(productKey As String) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function SetExtraActivationData(extraData As String) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function ActivateProduct() As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function DeactivateProduct() As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function ActivateProductOffline(filePath As String) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GenerateOfflineActivationRequest(filePath As String) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GenerateOfflineDeactivationRequest(filePath As String) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function IsProductGenuine() As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function IsProductActivated() As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GetExtraActivationData(extraData As StringBuilder, length As Integer) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GetCustomLicenseField(fieldId As String, fieldValue As StringBuilder, length As Integer) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GetProductKey(productKey As StringBuilder, length As Integer) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GetDaysLeftToExpiration(ByRef daysLeft As UInteger) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GetProductKeyExpiryDate(ByRef daysLeft As UInteger) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function SetTrialKey(trialKey As String) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function ActivateTrial() As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function IsTrialGenuine() As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function ExtendTrial(trialExtensionKey As String) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function InitializeTrial(trialLength As UInteger) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function GetTrialDaysLeft(ByRef daysLeft As UInteger, trialType As TrialType) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function SetDayIntervalForServerCheck(dayInterval As UInteger) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function SetGracePeriodForNetworkError(gracePeriod As UInteger) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function SetNetworkProxy(proxy As String) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function SetUserLock(userLock As Boolean) As Integer
            End Function

            <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
            Public Shared Function SetCryptlexHost(host As String) As Integer
            End Function

#If LA_ANY_CPU Then
			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "SetProductFile", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function SetProductFile_x64(filePath As String) As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "SetVersionGUID", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function SetVersionGUID_x64(versionGUID As String, flags As PermissionFlags) As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "SetProductKey", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function SetProductKey_x64(productKey As String) As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "SetExtraActivationData", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function SetExtraActivationData_x64(extraData As String) As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "ActivateProduct", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function ActivateProduct_x64() As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "DeactivateProduct", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function DeactivateProduct_x64() As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "ActivateProductOffline", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function ActivateProductOffline_x64(filePath As String) As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "GenerateOfflineActivationRequest", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function GenerateOfflineActivationRequest_x64(filePath As String) As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "GenerateOfflineDeactivationRequest", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function GenerateOfflineDeactivationRequest_x64(filePath As String) As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "IsProductGenuine", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function IsProductGenuine_x64() As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "IsProductActivated", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function IsProductActivated_x64() As Integer
			End Function

            <DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "GetExtraActivationData", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function GetExtraActivationData_x64(extraData As StringBuilder, length As Integer) As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "GetCustomLicenseField", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function GetCustomLicenseField_x64(fieldId As String, fieldValue As StringBuilder, length As Integer) As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "GetProductKey", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function GetProductKey_x64(productKey As StringBuilder, length As Integer) As Integer
			End Function

            <DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "GetDaysLeftToExpiration", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function GetDaysLeftToExpiration_x64(ByRef daysLeft As UInteger) As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "GetProductKeyExpiryDate", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function GetProductKeyExpiryDate_x64(ByRef expiryDate As UInteger) As Integer
			End Function

                        <DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "SetTrialKey", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function SetTrialKey_x64(trialKey As String) As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "ActivateTrial", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function ActivateTrial_x64() As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "IsTrialGenuine", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function IsTrialGenuine_x64() As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "ExtendTrial", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function ExtendTrial_x64(trialExtensionKey As String) As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "InitializeTrial", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function InitializeTrial_x64(trialLength As UInteger) As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "GetTrialDaysLeft", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function GetTrialDaysLeft_x64(ByRef daysLeft As UInteger, trialType As TrialType) As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "SetDayIntervalForServerCheck", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function SetDayIntervalForServerCheck_x64(dayInterval As UInteger) As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "SetGracePeriodForNetworkError", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function SetGracePeriodForNetworkError_x64(gracePeriod As UInteger) As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "SetNetworkProxy", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function SetNetworkProxy_x64(proxy As String) As Integer
			End Function

            <DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "SetUserLock", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function SetUserLock_x64(userLock As Boolean) As Integer
			End Function

			<DllImport(DLL_FILE_NAME_X64, CharSet := CharSet.Unicode, EntryPoint := "SetCryptlexHost", CallingConvention := CallingConvention.Cdecl)> _
			Public Shared Function SetCryptlexHost_x64(host As String) As Integer
			End Function
#End If
        End Class
    End Class
End Namespace

