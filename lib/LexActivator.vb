
Imports System.Runtime.InteropServices
Imports System.Text

Namespace Cryptlex
    Public NotInheritable Class LexActivator

        Private Sub New()
        End Sub

        Private Const DLL_FILE_NAME As String = "LexActivator.dll"

        '
        '     In order to use "Any CPU" configuration, rename 64 bit LexActivator.dll to LexActivator64.dll
        '     and add "LA_ANY_CPU" conditional compilation symbol.
        '        
        '     #Const LA_ANY_CPU = 1
        '

#If LA_ANY_CPU Then
        Private Const DLL_FILE_NAME_X64 As String = "LexActivator64.dll"
#End If

        Public Enum PermissionFlags As UInteger
            LA_USER = 1
            LA_SYSTEM = 2
        End Enum

        '
        '     FUNCTION: SetProductFile()

        '     PURPOSE: Sets the absolute path of the Product.dat file.

        '     This function must be called on every start of your program
        '     before any other functions are called.

        '     PARAMETERS:
        '     * filePath - absolute path of the product file (Product.dat)

        '     RETURN CODES: LA_OK, LA_E_FPATH, LA_E_PFILE

        '     NOTE: If this function fails to set the path of product file, none of the
        '     other functions will work.
        '

        Public Shared Function SetProductFile(filePath As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.SetProductFile_x64(filePath), Native.SetProductFile(filePath))
#Else
            Return Native.SetProductFile(filePath)
#End If

        End Function  

        '
        '     FUNCTION: SetProductData()

        '     PURPOSE: Embeds the Product.dat file in the application.

        '     It can be used instead of SetProductFile() in case you want
        '     to embed the Product.dat file in your application.

        '     This function must be called on every start of your program
        '     before any other functions are called.

        '     PARAMETERS:
        '     * productData - content of the Product.dat file

        '     RETURN CODES: LA_OK, LA_E_PDATA

        '     NOTE: If this function fails to set the product data, none of the
        '     other functions will work.
        '

        Public Shared Function SetProductData(productData As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.SetProductData_x64(productData), Native.SetProductData(productData))
#Else
            Return Native.SetProductData(productData)
#End If

        End Function 

        '
        '     FUNCTION: SetProductVersionGuid()

        '     PURPOSE: Sets the version GUID of your application.

        '     This function must be called on every start of your program before
        '     any other functions are called, with the exception of SetProductFile()
        '     or SetProductData() function.

        '     PARAMETERS:
        '     * versionGuid - the unique version GUID of your application as mentioned
        '     on the product version page of your application in the dashboard.

        '     * flags - depending upon whether your application requires admin/root
        '     permissions to run or not, this parameter can have one of the following
        '     values: LA_SYSTEM, LA_USER

        '     RETURN CODES: LA_OK, LA_E_WMIC, LA_E_PFILE, LA_E_PDATA, LA_E_GUID, LA_E_PERMISSION

        '     NOTE: If this function fails to set the version GUID, none of the other
        '     functions will work.
        '

        Public Shared Function SetProductVersionGuid(versionGuid As String, flags As PermissionFlags) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.SetProductVersionGuid_x64(versionGuid, flags), Native.SetProductVersionGuid(versionGuid, flags))
#Else
            Return Native.SetProductVersionGuid(versionGuid, flags)
#End If

        End Function

        '
        '     FUNCTION: SetProductKey()

        '     PURPOSE: Sets the product key required to activate the application.

        '     PARAMETERS:
        '     * productKey - a valid product key generated for the application.

        '     RETURN CODES: LA_OK, LA_E_GUID, LA_E_PKEY
        '

        Public Shared Function SetProductKey(productKey As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.SetProductKey_x64(productKey), Native.SetProductKey(productKey))
#Else
            Return Native.SetProductKey(productKey)
#End If

        End Function

        '
        '     FUNCTION: SetActivationExtraData()

        '     PURPOSE: Sets the extra data which you may want to fetch from the user
        '     at the time of activation.

        '     The extra data appears along with the activation details of the product key
        '     in dashboard.

        '     PARAMETERS:
        '     * extraData - string of maximum length 1024 characters with utf-8 encoding.

        '     RETURN CODES: LA_OK, LA_E_GUID, LA_E_EDATA_LEN
        '

        Public Shared Function SetActivationExtraData(extraData As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.SetActivationExtraData_x64(extraData), Native.SetActivationExtraData(extraData))
#Else
            Return Native.SetActivationExtraData(extraData)
#End If

        End Function

        '
        '     FUNCTION: SetTrialActivationExtraData()

        '     PURPOSE: Sets the extra data which you may want to fetch from the user
        '     at the time of trial activation.

        '     The extra data appears along with the trial activation details in dashboard.

        '     PARAMETERS:
        '     * extraData - string of maximum length 1024 characters with utf-8 encoding.

        '     RETURN CODES: LA_OK, LA_E_GUID, LA_E_EDATA_LEN
        '

        Public Shared Function SetTrialActivationExtraData(extraData As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.SetTrialActivationExtraData_x64(extraData), Native.SetTrialActivationExtraData(extraData))
#Else
            Return Native.SetTrialActivationExtraData(extraData)
#End If

        End Function

        '
        '     FUNCTION: SetNetworkProxy()

        '     PURPOSE: Sets the network proxy to be used when contacting Cryptlex servers.

        '     The proxy format should be: [protocol://][username:password@]machine[:port]

        '     Following are some examples of the valid proxy strings:
        '         - http://127.0.0.1:8000/
        '         - http://user:pass@127.0.0.1:8000/
        '         - socks5://127.0.0.1:8000/

        '     PARAMETERS:
        '     * proxy - proxy string having correct proxy format

        '     RETURN CODES: LA_OK, LA_E_GUID, LA_E_NET_PROXY

        '     NOTE: Proxy settings of the computer are automatically detected. So, in most of the
        '     cases you don't need to care whether your user is behind a proxy server or not.
        '

        Public Shared Function SetNetworkProxy(proxy As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.SetNetworkProxy_x64(proxy), Native.SetNetworkProxy(proxy))
#Else
            Return Native.SetNetworkProxy(proxy)
#End If
        End Function

        '
        '     FUNCTION: GetAppVersion()

        '     PURPOSE: Gets the app version of the product as set in the dashboard.

        '     PARAMETERS:
        '     * appVersion - pointer to a buffer that receives the value of the string
        '     * length - size of the buffer pointed to by the appVersion parameter

        '     RETURN CODES: LA_OK, LA_E_GUID, LA_FAIL, LA_E_TIME, LA_E_BUFFER_SIZE
        '

        Public Shared Function GetAppVersion(appVersion As StringBuilder, length As Integer) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GetAppVersion_x64(appVersion, length), Native.GetAppVersion(appVersion, length))
#Else
            Return Native.GetAppVersion(appVersion, length)
#End If
        End Function

        '
        '     FUNCTION: GetProductKey()

        '     PURPOSE: Gets the stored product key which was used for activation.

        '     PARAMETERS:
        '     * productKey - pointer to a buffer that receives the value of the string
        '     * length - size of the buffer pointed to by the productKey parameter

        '     RETURN CODES: LA_OK, LA_E_PKEY, LA_E_GUID, LA_E_BUFFER_SIZE
        '

        Public Shared Function GetProductKey(productKey As StringBuilder, length As Integer) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GetProductKey_x64(productKey, length), Native.GetProductKey(productKey, length))
#Else
            Return Native.GetProductKey(productKey, length)
#End If
        End Function

        '
        '     FUNCTION: GetProductKeyEmail()

        '     PURPOSE: Gets the email associated with product key used for activation.

        '     PARAMETERS:
        '     * productKey - pointer to a buffer that receives the value of the string
        '     * length - size of the buffer pointed to by the productKeyEmail parameter

        '     RETURN CODES: LA_OK, LA_E_GUID, LA_FAIL, LA_E_TIME, LA_E_BUFFER_SIZE
        '

        Public Shared Function GetProductKeyEmail(productKeyEmail As StringBuilder, length As Integer) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GetProductKeyEmail_x64(productKeyEmail, length), Native.GetProductKeyEmail(productKeyEmail, length))
#Else
            Return Native.GetProductKeyEmail(productKeyEmail, length)
#End If
        End Function

        '
        '     FUNCTION: GetProductKeyExpiryDate()

        '     PURPOSE: Gets the product key expiry date timestamp.

        '     PARAMETERS:
        '     * expiryDate - pointer to the integer that receives the value

        '     RETURN CODES: LA_OK, LA_E_GUID, LA_FAIL, LA_E_TIME
        '

        Public Shared Function GetProductKeyExpiryDate(ByRef expiryDate As UInteger) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GetProductKeyExpiryDate_x64(expiryDate), Native.GetProductKeyExpiryDate(expiryDate))
#Else
            Return Native.GetProductKeyExpiryDate(expiryDate)
#End If
        End Function

        '
        '     FUNCTION: GetProductKeyCustomField()

        '     PURPOSE: Get the value of the custom field associated with the product key.

        '     PARAMETERS:
        '     * fieldId - id of the custom field whose value you want to get
        '     * fieldValue - pointer to a buffer that receives the value of the string
        '     * length - size of the buffer pointed to by the fieldValue parameter

        '     RETURN CODES: LA_OK, LA_E_GUID, LA_FAIL, LA_E_TIME, LA_E_CUSTOM_FIELD_ID,
        '     LA_E_BUFFER_SIZE
        '

        Public Shared Function GetProductKeyCustomField(fieldId As String, fieldValue As StringBuilder, length As Integer) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GetProductKeyCustomField_x64(fieldId, fieldValue, length), Native.GetProductKeyCustomField(fieldId, fieldValue, length))
#Else
            Return Native.GetProductKeyCustomField(fieldId, fieldValue, length)
#End If
        End Function

        '
        '     FUNCTION: GetActivationExtraData()

        '     PURPOSE: Gets the value of the activation extra data.

        '     PARAMETERS:
        '     * extraData - pointer to a buffer that receives the value of the string
        '     * length - size of the buffer pointed to by the fieldValue parameter

        '     RETURN CODES: LA_OK, LA_E_GUID, LA_FAIL, LA_E_TIME, LA_E_BUFFER_SIZE
        '

        Public Shared Function GetActivationExtraData(extraData As StringBuilder, length As Integer) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GetActivationExtraData_x64(extraData, length), Native.GetActivationExtraData(extraData, length))
#Else
            Return Native.GetActivationExtraData(extraData, length)
#End If
        End Function

        '
        '     FUNCTION: GetTrialActivationExtraData()

        '     PURPOSE: Gets the value of the trial activation extra data.

        '     PARAMETERS:
        '     * extraData - pointer to a buffer that receives the value of the string
        '     * length - size of the buffer pointed to by the fieldValue parameter

        '     RETURN CODES: LA_OK, LA_E_GUID, LA_FAIL, LA_E_TIME, LA_E_BUFFER_SIZE
        '

        Public Shared Function GetTrialActivationExtraData(extraData As StringBuilder, length As Integer) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GetTrialActivationExtraData_x64(extraData, length), Native.GetTrialActivationExtraData(extraData, length))
#Else
            Return Native.GetTrialActivationExtraData(extraData, length)
#End If
        End Function

        '
        '     FUNCTION: GetTrialExpiryDate()

        '     PURPOSE: Gets the trial expiry date timestamp.

        '     PARAMETERS:
        '     * trialExpiryDate - pointer to the integer that receives the value

        '     RETURN CODES: LA_OK, LA_E_GUID, LA_FAIL, LA_E_TIME
        '


        Public Shared Function GetTrialExpiryDate(ByRef trialExpiryDate As UInteger) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GetTrialExpiryDate_x64(trialExpiryDate), Native.GetTrialExpiryDate(trialExpiryDate))
#Else
            Return Native.GetTrialExpiryDate(trialExpiryDate)
#End If
        End Function

        '
        '     FUNCTION: GetLocalTrialExpiryDate()

        '     PURPOSE: Gets the trial expiry date timestamp.

        '     PARAMETERS:
        '     * trialExpiryDate - pointer to the integer that receives the value

        '     RETURN CODES: LA_OK, LA_E_GUID, LA_FAIL, LA_E_TIME
        '

        Public Shared Function GetLocalTrialExpiryDate(ByRef trialExpiryDate As UInteger) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GetLocalTrialExpiryDate_x64(trialExpiryDate), Native.GetLocalTrialExpiryDate(trialExpiryDate))
#Else
            Return Native.GetLocalTrialExpiryDate(trialExpiryDate)
#End If
        End Function

        '
        '     FUNCTION: ActivateProduct()

        '     PURPOSE: Activates your application by contacting the Cryptlex servers. It
        '     validates the key and returns with encrypted and digitally signed token
        '     which it stores and uses to activate your application.

        '     This function should be executed at the time of registration, ideally on
        '     a button click.

        '     RETURN CODES: LA_OK, LA_EXPIRED, LA_REVOKED, LA_FAIL, LA_E_GUID, LA_E_PKEY,
        '     LA_E_INET, LA_E_VM, LA_E_TIME, LA_E_ACT_LIMIT, LA_E_SERVER, LA_E_CLIENT,
        '     LA_E_PKEY_TYPE, LA_E_COUNTRY, LA_E_IP
        '

        Public Shared Function ActivateProduct() As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.ActivateProduct_x64(), Native.ActivateProduct())
#Else
            Return Native.ActivateProduct()
#End If

        End Function

        '
        '     FUNCTION: ActivateProductOffline()

        '     PURPOSE: Activates your application using the offline activation response
        '     file.

        '     PARAMETERS:
        '     * filePath - path of the offline activation response file.

        '     RETURN CODES: LA_OK, LA_EXPIRED, LA_FAIL, LA_E_GUID, LA_E_PKEY, LA_E_OFILE
        '     LA_E_VM, LA_E_TIME, LA_E_FPATH, LA_E_OFILE_EXPIRED
        '

        Public Shared Function ActivateProductOffline(filePath As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.ActivateProductOffline_x64(filePath), Native.ActivateProductOffline(filePath))
#Else
            Return Native.ActivateProductOffline(filePath)
#End If
        End Function

        '
        '     FUNCTION: GenerateOfflineActivationRequest()

        '     PURPOSE: Generates the offline activation request needed for generating
        '     offline activation response in the dashboard.

        '     PARAMETERS:
        '     * filePath - path of the file for the offline request.

        '     RETURN CODES: LA_OK, LA_FAIL, LA_E_GUID, LA_E_PKEY, LA_E_FILE_PERMISSION
        '

        Public Shared Function GenerateOfflineActivationRequest(filePath As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GenerateOfflineActivationRequest_x64(filePath), Native.GenerateOfflineActivationRequest(filePath))
#Else
            Return Native.GenerateOfflineActivationRequest(filePath)
#End If
        End Function

        '
        '     FUNCTION: DeactivateProduct()

        '     PURPOSE: Deactivates the application and frees up the corresponding activation
        '     slot by contacting the Cryptlex servers.

        '     This function should be executed at the time of de-registration, ideally on
        '     a button click.

        '     RETURN CODES: LA_OK, LA_E_DEACT_LIMIT, LA_FAIL, LA_E_GUID, LA_E_TIME
        '     LA_E_PKEY, LA_E_INET, LA_E_SERVER, LA_E_CLIENT
        '

        Public Shared Function DeactivateProduct() As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.DeactivateProduct_x64(), Native.DeactivateProduct())
#Else
            Return Native.DeactivateProduct()
#End If
        End Function

        '
        '     FUNCTION: GenerateOfflineDeactivationRequest()

        '     PURPOSE: Generates the offline deactivation request needed for deactivation of
        '     the product key in the dashboard and deactivates the application.

        '     A valid offline deactivation file confirms that the application has been successfully
        '     deactivated on the user's machine.

        '     PARAMETERS:
        '     * filePath - path of the file for the offline request.

        '     RETURN CODES: LA_OK, LA_FAIL, LA_E_GUID, LA_E_PKEY, LA_E_FILE_PERMISSION,
        '     LA_E_TIME
        '

        Public Shared Function GenerateOfflineDeactivationRequest(filePath As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.GenerateOfflineDeactivationRequest_x64(filePath), Native.GenerateOfflineDeactivationRequest(filePath))
#Else
            Return Native.GenerateOfflineDeactivationRequest(filePath)
#End If
        End Function

        '
        '     FUNCTION: IsProductGenuine()

        '     PURPOSE: It verifies whether your app is genuinely activated or not. The verification is
        '     done locally by verifying the cryptographic digital signature fetched at the time of
        '     activation.

        '     After verifying locally, it schedules a server check in a separate thread. After the
        '     first server sync it periodically does further syncs at a frequency set for the product
        '     key.

        '     In case server sync fails due to network error, and it continues to fail for fixed
        '     number of days (grace period), the function returns LA_GP_OVER instead of LA_OK.

        '     This function must be called on every start of your program to verify the activation
        '     of your app.

        '     RETURN CODES: LA_OK, LA_EXPIRED, LA_REVOKED, LA_GP_OVER, LA_FAIL, LA_E_GUID, LA_E_PKEY,
        '     LA_E_TIME

        '     NOTE: If application was activated offline using ActivateProductOffline() function, you
        '     may want to set grace period to 0 to ignore grace period.
        '

        Public Shared Function IsProductGenuine() As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.IsProductGenuine_x64(), Native.IsProductGenuine())
#Else
            Return Native.IsProductGenuine()
#End If
        End Function

        '
        '     FUNCTION: IsProductActivated()

        '     PURPOSE: It verifies whether your app is genuinely activated or not. The verification is
        '     done locally by verifying the cryptographic digital signature fetched at the time of
        '     activation.

        '     This is just an auxiliary function which you may use in some specific cases, when you
        '     want to skip the server sync.

        '     RETURN CODES: LA_OK, LA_EXPIRED, LA_REVOKED, LA_GP_OVER, LA_FAIL, LA_E_GUID, LA_E_PKEY,
        '     LA_E_TIME

        '     NOTE: You may want to set grace period to 0 to ignore grace period.
        '

        Public Shared Function IsProductActivated() As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.IsProductActivated_x64(), Native.IsProductActivated())
#Else
            Return Native.IsProductActivated()
#End If
        End Function

        '
        '     FUNCTION: ActivateTrial()

        '     PURPOSE: Starts the verified trial in your application by contacting the
        '     Cryptlex servers.

        '     This function should be executed when your application starts first time on
        '     the user's computer, ideally on a button click.

        '     RETURN CODES: LA_OK, LA_T_EXPIRED, LA_FAIL, LA_E_GUID, LA_E_INET,
        '     LA_E_VM, LA_E_TIME, LA_E_SERVER, LA_E_CLIENT, LA_E_COUNTRY, LA_E_IP
        '

        Public Shared Function ActivateTrial() As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.ActivateTrial_x64(), Native.ActivateTrial())
#Else
            Return Native.ActivateTrial()
#End If
        End Function

        '
        '     FUNCTION: IsTrialGenuine()

        '     PURPOSE: It verifies whether trial has started and is genuine or not. The
        '     verification is done locally by verifying the cryptographic digital signature
        '     fetched at the time of trial activation.

        '     This function must be called on every start of your program during the trial period.

        '     RETURN CODES: LA_OK, LA_T_EXPIRED, LA_FAIL, LA_E_TIME, LA_E_GUID

        '

        Public Shared Function IsTrialGenuine() As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.IsTrialGenuine_x64(), Native.IsTrialGenuine())
#Else
            Return Native.IsTrialGenuine()
#End If
        End Function

        '
        '     FUNCTION: ExtendTrial()

        '     PURPOSE: Extends the trial using the trial extension key generated in the dashboard
        '     for the product version.

        '     PARAMETERS:
        '     * trialExtensionKey - trial extension key generated for the product version

        '     RETURN CODES: LA_OK, LA_T_EXPIRED, LA_FAIL, LA_E_GUID, LA_E_INET,
        '     LA_E_VM, LA_E_TIME, LA_E_TEXT_KEY, LA_E_SERVER, LA_E_CLIENT,
        '     LA_E_TRIAL_NOT_EXPIRED

        '     NOTE: The function is only meant for verified trials.
        '

        Public Shared Function ExtendTrial(trialExtensionKey As String) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.ExtendTrial_x64(trialExtensionKey), Native.ExtendTrial(trialExtensionKey))
#Else
            Return Native.ExtendTrial(trialExtensionKey)
#End If
        End Function
        
        '
        '     FUNCTION: ActivateLocalTrial()

        '     PURPOSE: Starts the local(unverified) trial.

        '     This function should be executed when your application starts first time on
        '     the user's computer, ideally on a button click.

        '     PARAMETERS:
        '     * trialLength - trial length as set in the dashboard for the product version

        '     RETURN CODES: LA_OK, LA_LT_EXPIRED, LA_FAIL, LA_E_GUID, LA_E_TIME

        '     NOTE: The function is only meant for unverified trials.
        '

        Public Shared Function ActivateLocalTrial(trialLength As UInteger) As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.ActivateLocalTrial_x64(trialLength), Native.ActivateLocalTrial(trialLength))
#Else
            Return Native.ActivateLocalTrial(trialLength)
#End If
        End Function

        '
        '     FUNCTION: IsLocalTrialGenuine()

        '     PURPOSE: It verifies whether trial has started and is genuine or not. The
        '     verification is done locally.

        '     This function must be called on every start of your program during the trial period.

        '     RETURN CODES: LA_OK, LA_LT_EXPIRED, LA_FAIL, LA_E_GUID, LA_E_TIME

        '     NOTE: The function is only meant for unverified trials.
        '

        Public Shared Function IsLocalTrialGenuine() As Integer
#If LA_ANY_CPU Then
			Return If(IntPtr.Size = 8, Native.IsLocalTrialGenuine_x64(), Native.IsLocalTrialGenuine())
#Else
            Return Native.IsLocalTrialGenuine()
#End If
        End Function

        '** Return Codes **


        Public Const LA_OK As Integer = 0

        Public Const LA_FAIL As Integer = 1


        '
        ' CODE: LA_EXPIRED

        ' MESSAGE: The product key has expired or system time has been tampered
        ' with. Ensure your date and time settings are correct.
        '

        Public Const LA_EXPIRED       As Integer = 2

        '
        ' CODE: LA_REVOKED

        ' MESSAGE: The product key has been revoked.
        '

        Public Const LA_REVOKED       As Integer = 3

        '
        ' CODE: LA_GP_OVER

        ' MESSAGE: The grace period is over.
        '

        Public Const LA_GP_OVER       As Integer = 4

        '
        ' CODE: LA_E_INET

        ' MESSAGE: Failed to connect to the server due to network error.
        '

        Public Const LA_E_INET		 As Integer = 5

        '
        ' CODE: LA_E_PKEY

        ' MESSAGE: Invalid product key.
        '

        Public Const LA_E_PKEY		 As Integer = 6

        '
        ' CODE: LA_E_PFILE

        ' MESSAGE: Invalid or corrupted product file.
        '

        Public Const LA_E_PFILE		 As Integer = 7

        '
        ' CODE: LA_E_FPATH

        ' MESSAGE: Invalid product file path.
        '

        Public Const LA_E_FPATH		 As Integer = 8

        '
        ' CODE: LA_E_GUID

        ' MESSAGE: The version GUID doesn't match that of the product file.
        '

        Public Const LA_E_GUID		 As Integer = 9

        '
        ' CODE: LA_E_OFILE

        ' MESSAGE: Invalid offline activation response file.
        '

        Public Const LA_E_OFILE		 As Integer = 10

        '
        ' CODE: LA_E_PERMISSION

        ' MESSAGE: Insufficent system permissions. Occurs when LA_SYSTEM flag is used
        ' but application is not run with admin privileges.
        '

        Public Const LA_E_PERMISSION  As Integer = 11

        '
        ' CODE: LA_E_EDATA_LEN

        ' MESSAGE: Extra activation data length is more than 256 characters.
        '

        Public Const LA_E_EDATA_LEN   As Integer = 12

        '
        ' CODE: LA_E_PKEY_TYPE

        ' MESSAGE: Invalid product key type.
        '

        Public Const LA_E_PKEY_TYPE   As Integer = 13

        '
        ' CODE: LA_E_TIME

        ' MESSAGE: The system time has been tampered with. Ensure your date
        ' and time settings are correct.
        '

        Public Const LA_E_TIME        As Integer = 14

        '
        ' CODE: LA_E_VM

        ' MESSAGE: Application is being run inside a virtual machine / hypervisor,
        ' and activation has been disallowed in the VM.
        ' but
        '

        Public Const LA_E_VM          As Integer = 15

        '
        ' CODE: LA_E_WMIC

        ' MESSAGE: Fingerprint couldn't be generated because Windows Management 
        ' Instrumentation (WMI service has been disabled. This error is specific
        ' to Windows only.
        '

        Public Const LA_E_WMIC        As Integer = 16

        '
        ' CODE: LA_E_TEXT_KEY

        ' MESSAGE: Invalid trial extension key.
        '

        Public Const LA_E_TEXT_KEY    As Integer = 17

        '
        ' CODE: LA_E_OFILE_EXPIRED

        ' MESSAGE: The offline activation response has expired.
        '

        Public Const LA_E_OFILE_EXPIRED   As Integer = 18

        '
        ' CODE: LA_T_EXPIRED

        ' MESSAGE: The trial has expired or system time has been tampered
        ' with. Ensure your date and time settings are correct.
        '

        Public Const LA_T_EXPIRED      As Integer = 19

        '
        ' CODE: LA_LT_EXPIRED

        ' MESSAGE: The local trial has expired or system time has been tampered
        ' with. Ensure your date and time settings are correct.
        '

        Public Const LA_LT_EXPIRED     As Integer = 20

        '
        ' CODE: LA_E_BUFFER_SIZE

        ' MESSAGE: The buffer size was smaller than required.
        '

        Public Const LA_E_BUFFER_SIZE   As Integer = 21

        '
        ' CODE: LA_E_CUSTOM_FIELD_ID

        ' MESSAGE: Invalid custom field id.
        '

        Public Const LA_E_CUSTOM_FIELD_ID  As Integer = 22

        '
        ' CODE: LA_E_NET_PROXY

        ' MESSAGE: Invalid network proxy.
        '

        Public Const LA_E_NET_PROXY   As Integer = 23

        '
        ' CODE: LA_E_HOST_URL

        ' MESSAGE: Invalid Cryptlex host url.
        '

        Public Const LA_E_HOST_URL  As Integer = 24

        '
        ' CODE: LA_E_DEACT_LIMIT

        ' MESSAGE: Deactivation limit for key has reached
        '

        Public Const LA_E_DEACT_LIMIT   As Integer = 25

        '
        ' CODE: LA_E_ACT_LIMIT

        ' MESSAGE: Activation limit for key has reached
        '

        Public Const LA_E_ACT_LIMIT    As Integer = 26

        '
        ' CODE: LA_E_PDATA

        ' MESSAGE: Invalid product data
        '

        Public Const LA_E_PDATA    As Integer = 27

        '
        ' CODE: LA_E_TRIAL_NOT_EXPIRED

        ' MESSAGE: Trial has not expired.
        '

        Public Const LA_E_TRIAL_NOT_EXPIRED  As Integer = 28

        '
        ' CODE: LA_E_COUNTRY

        ' MESSAGE: Country is not allowed
        '

        Public Const LA_E_COUNTRY   As Integer = 29

        '
        ' CODE: LA_E_IP

        ' MESSAGE: IP address is not allowed
        '

        Public Const LA_E_IP   As Integer = 30

        '
        ' CODE: LA_E_FILE_PERMISSION

        ' MESSAGE: No permission to write to file
        '

        Public Const LA_E_FILE_PERMISSION  As Integer = 31

        '
        ' CODE: LA_E_SERVER

        ' MESSAGE: Server error
        '

        Public Const LA_E_SERVER  As Integer = 32

        '
        ' CODE: LA_E_CLIENT

        ' MESSAGE: Client error
        '

        Public Const LA_E_CLIENT     As Integer = 33


        Private NotInheritable Class Native

                Private Sub New()
                End Sub

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function SetProductFile(filePath As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function SetProductData(productData As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function SetProductVersionGuid(versionGuid As String, flags As PermissionFlags) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function SetProductKey(productKey As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function SetActivationExtraData(extraData As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function SetTrialActivationExtraData(extraData As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function SetNetworkProxy(proxy As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GetAppVersion(appVersion As StringBuilder, length As Integer) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GetProductKey(productKey As StringBuilder, length As Integer) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GetProductKeyEmail(productKeyEmail As StringBuilder, length As Integer) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GetProductKeyExpiryDate(ByRef expiryDate As UInteger) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GetProductKeyCustomField(fieldId As String, fieldValue As StringBuilder, length As Integer) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GetActivationExtraData(extraData As StringBuilder, length As Integer) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GetTrialActivationExtraData(extraData As StringBuilder, length As Integer) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GetTrialExpiryDate(ByRef trialExpiryDate As UInteger) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GetLocalTrialExpiryDate(ByRef trialExpiryDate As UInteger) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function ActivateProduct() As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function ActivateProductOffline(filePath As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GenerateOfflineActivationRequest(filePath As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function DeactivateProduct() As Integer
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
                Public Shared Function ActivateTrial() As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function IsTrialGenuine() As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function ExtendTrial(trialExtensionKey As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function ActivateLocalTrial(trialLength As UInteger) As Integer
                End Function

                <DllImport(DLL_FILE_NAME, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function IsLocalTrialGenuine() As Integer
                End Function

#If LA_ANY_CPU Then

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="SetProductFile", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function SetProductFile_x64(filePath As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="SetProductData", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function SetProductData_x64(productData As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="SetProductVersionGuid", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function SetProductVersionGuid_x64(versionGuid As String, flags As PermissionFlags) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="SetProductKey", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function SetProductKey_x64(productKey As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="SetActivationExtraData", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function SetActivationExtraData_x64(extraData As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="SetTrialActivationExtraData", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function SetTrialActivationExtraData_x64(extraData As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="SetNetworkProxy", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function SetNetworkProxy_x64(proxy As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="GetAppVersion", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GetAppVersion_x64(appVersion As StringBuilder, length As Integer) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="GetProductKey", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GetProductKey_x64(productKey As StringBuilder, length As Integer) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="GetProductKeyEmail", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GetProductKeyEmail_x64(productKeyEmail As StringBuilder, length As Integer) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="GetProductKeyExpiryDate", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GetProductKeyExpiryDate_x64(ByRef expiryDate As UInteger) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="GetProductKeyCustomField", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GetProductKeyCustomField_x64(fieldId As String, fieldValue As StringBuilder, length As Integer) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="GetActivationExtraData", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GetActivationExtraData_x64(extraData As StringBuilder, length As Integer) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="GetTrialActivationExtraData", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GetTrialActivationExtraData_x64(extraData As StringBuilder, length As Integer) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="GetTrialExpiryDate", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GetTrialExpiryDate_x64(ByRef trialExpiryDate As UInteger) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="GetLocalTrialExpiryDate", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GetLocalTrialExpiryDate_x64(ByRef trialExpiryDate As UInteger) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="ActivateProduct", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function ActivateProduct_x64() As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="ActivateProductOffline", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function ActivateProductOffline_x64(filePath As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="GenerateOfflineActivationRequest", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GenerateOfflineActivationRequest_x64(filePath As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="DeactivateProduct", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function DeactivateProduct_x64() As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="GenerateOfflineDeactivationRequest", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function GenerateOfflineDeactivationRequest_x64(filePath As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="IsProductGenuine", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function IsProductGenuine_x64() As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="IsProductActivated", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function IsProductActivated_x64() As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="ActivateTrial", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function ActivateTrial_x64() As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="IsTrialGenuine", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function IsTrialGenuine_x64() As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="ExtendTrial", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function ExtendTrial_x64(trialExtensionKey As String) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="ActivateLocalTrial", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function ActivateLocalTrial_x64(trialLength As UInteger) As Integer
                End Function

                <DllImport(DLL_FILE_NAME_X64, CharSet:=CharSet.Unicode, EntryPoint:="IsLocalTrialGenuine", CallingConvention:=CallingConvention.Cdecl)>
                Public Shared Function IsLocalTrialGenuine_x64() As Integer
                End Function

#End If
        End Class
    End Class
End Namespace

