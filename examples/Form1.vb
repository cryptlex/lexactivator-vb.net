Imports System.IO
Imports Sample.Cryptlex

Public Class Form1
    Public Sub New()

        InitializeComponent()
        Dim status As Integer
        ' status = LexActivator.SetProductFile("ABSOLUTE_PATH_OF_PRODUCT.DAT_FILE")
        status = LexActivator.SetProductData("PASTE_CONTENT_OF_PRODUCT.DAT_FILE")
        If status <> LexActivator.LA_OK Then
            Me.statusLabel.Text = "Error setting product file: " & status.ToString()
            Return
        End If

        status = LexActivator.SetProductVersionGuid("PASTE_PRODUCT_VERSION_GUID", LexActivator.PermissionFlags.LA_USER)
        If status <> LexActivator.LA_OK Then
            Me.statusLabel.Text = "Error setting version GUID: " & status.ToString()
            Return
        End If

        status = LexActivator.IsProductGenuine()
        If status = LexActivator.LA_OK OrElse status = LexActivator.LA_EXPIRED OrElse status = LexActivator.LA_REVOKED OrElse status = LexActivator.LA_GP_OVER Then
            ' Dim expiryDate As UInteger = 0
            ' LexActivator.GetProductKeyExpiryDate(expiryDate)
            ' Dim daysLeft As Integer = CInt((expiryDate - unixTimestamp())) / 86500
            Me.statusLabel.Text = "Product genuinely activated! Activation Status: " & status.ToString()
            Me.activateBtn.Text = "Deactivate"
            Me.activateTrialBtn.Enabled = False
            Return
        End If

        status = LexActivator.IsTrialGenuine()
        If status = LexActivator.LA_OK Then
            Dim trialExpiryDate As UInteger = 0
            LexActivator.GetTrialExpiryDate(trialExpiryDate)
            Dim daysLeft As Integer = CInt((trialExpiryDate - unixTimestamp())) / 86500
            Me.statusLabel.Text = "Trial period! Days left:" & daysLeft.ToString()
            Me.activateTrialBtn.Enabled = False
        ElseIf status = LexActivator.LA_T_EXPIRED Then
            Me.statusLabel.Text = "Trial has expired!"
        Else
            Me.statusLabel.Text = "Trial has not started or has been tampered: " & status.ToString()
        End If

    End Sub
    Public Sub activateBtn_Click(sender As Object, e As EventArgs) Handles activateBtn.Click
        Dim status As Integer
        If Me.activateBtn.Text = "Deactivate" Then
            status = LexActivator.DeactivateProduct()
            If status = LexActivator.LA_OK Then
                Me.statusLabel.Text = "Product deactivated successfully"
                Me.activateBtn.Text = "Activate"
                Me.activateTrialBtn.Enabled = True
                Return
            End If

            Me.statusLabel.Text = "Error deactivating product: " & status.ToString()
            Return
        End If

        status = LexActivator.SetProductKey(productKeyBox.Text)
        If status <> LexActivator.LA_OK Then
            Me.statusLabel.Text = "Error setting product key: " & status.ToString()
            Return
        End If

        status = LexActivator.SetActivationExtraData("SAMPLE DATA")
        If status <> LexActivator.LA_OK Then
            Me.statusLabel.Text = "Error setting activation extra data: " & status.ToString()
            Return
        End If

        status = LexActivator.ActivateProduct()
        If status <> LexActivator.LA_OK Then
            Me.statusLabel.Text = "Error activating the product: " & status.ToString()
            Return
        Else
            Me.statusLabel.Text = "Activation Successful"
            Me.activateBtn.Text = "Deactivate"
            Me.activateTrialBtn.Enabled = False
        End If
    End Sub

    Public Sub activateTrialBtn_Click(sender As Object, e As EventArgs) Handles activateTrialBtn.Click
        Dim status As Integer
        status = LexActivator.SetTrialActivationExtraData("SAMPLE DATA")
        If status <> LexActivator.LA_OK Then
            Me.statusLabel.Text = "Error setting activation extra data: " & status.ToString()
            Return
        End If

        status = LexActivator.ActivateTrial()
        If status <> LexActivator.LA_OK Then
            Me.statusLabel.Text = "Error activating the trial: " & status.ToString()
            Return
        Else
            Me.statusLabel.Text = "Trial started Successful"
        End If
    End Sub

    Private Function unixTimestamp() As UInteger
        Return CUInt((DateTime.UtcNow.Subtract(New DateTime(1970, 1, 1))).TotalSeconds)
    End Function

End Class
