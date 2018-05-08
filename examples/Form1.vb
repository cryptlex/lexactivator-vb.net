﻿Imports System.IO
Imports Sample.Cryptlex

Public Class Form1
    Public Sub New()

        InitializeComponent()
        Dim status As Integer
        status = LexActivator.SetProductData("PASTE_CONTENT_OF_PRODUCT.DAT_FILE")
        If status <> LexActivator.StatusCodes.LA_OK Then
            Me.statusLabel.Text = "Error setting product file: " & status.ToString()
            Return
        End If

        status = LexActivator.SetProductId("PASTE_PRODUCT_ID", LexActivator.PermissionFlags.LA_USER)
        If status <> LexActivator.StatusCodes.LA_OK Then
            Me.statusLabel.Text = "Error setting product id: " & status.ToString()
            Return
        End If

        status = LexActivator.IsLicenseGenuine()
        If status = LexActivator.StatusCodes.LA_OK OrElse status = LexActivator.StatusCodes.LA_EXPIRED OrElse status = LexActivator.StatusCodes.LA_SUSPENDED OrElse status = LexActivator.StatusCodes.LA_GRACE_PERIOD_OVER Then
            Me.statusLabel.Text = "License genuinely activated! Activation Status: " & status.ToString()
            Me.activateBtn.Text = "Deactivate"
            Me.activateTrialBtn.Enabled = False
            Return
        End If

        status = LexActivator.IsTrialGenuine()
        If status = LexActivator.StatusCodes.LA_OK Then
            Dim trialExpiryDate As UInteger = 0
            LexActivator.GetTrialExpiryDate(trialExpiryDate)
            Dim daysLeft As Integer = CInt((trialExpiryDate - unixTimestamp())) / 86500
            Me.statusLabel.Text = "Trial period! Days left:" & daysLeft.ToString()
            Me.activateTrialBtn.Enabled = False
        ElseIf status = LexActivator.StatusCodes.LA_TRIAL_EXPIRED Then
            Me.statusLabel.Text = "Trial has expired!"
        Else
            Me.statusLabel.Text = "Trial has not started or has been tampered: " & status.ToString()
        End If

    End Sub

    Public Sub activateBtn_Click(sender As Object, e As EventArgs) Handles activateBtn.Click
        Dim status As Integer
        If Me.activateBtn.Text = "Deactivate" Then
            status = LexActivator.DeactivateLicense()
            If status = LexActivator.StatusCodes.LA_OK Then
                Me.statusLabel.Text = "License deactivated successfully"
                Me.activateBtn.Text = "Activate"
                Me.activateTrialBtn.Enabled = True
                Return
            End If

            Me.statusLabel.Text = "Error deactivating license: " & status.ToString()
            Return
        End If

        status = LexActivator.SetLicenseKey(productKeyBox.Text)
        If status <> LexActivator.StatusCodes.LA_OK Then
            Me.statusLabel.Text = "Error setting license key: " & status.ToString()
            Return
        End If

        status = LexActivator.SetActivationMetadata("key1", "value1")
        If status <> LexActivator.StatusCodes.LA_OK Then
            Me.statusLabel.Text = "Error setting activation metadata: " & status.ToString()
            Return
        End If
        
        status = LexActivator.ActivateLicense()
        If status = LexActivator.StatusCodes.LA_OK OrElse status = LexActivator.StatusCodes.LA_EXPIRED OrElse status = LexActivator.StatusCodes.LA_SUSPENDED Then
            Me.statusLabel.Text = "Activation Successful: " & status.ToString()
            Me.activateBtn.Text = "Deactivate"
            Me.activateTrialBtn.Enabled = False
        Else
            Me.statusLabel.Text = "Error activating the license: " & status.ToString()
            Return
        End If
    End Sub

    Public Sub activateTrialBtn_Click(sender As Object, e As EventArgs) Handles activateTrialBtn.Click
        Dim status As Integer
        status = LexActivator.SetTrialActivationMetadata("key2", "value2")
        If status <> LexActivator.StatusCodes.LA_OK Then
            Me.statusLabel.Text = "Error setting activation metadata: " & status.ToString()
            Return
        End If

        status = LexActivator.ActivateTrial()
        If status <> LexActivator.StatusCodes.LA_OK Then
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
