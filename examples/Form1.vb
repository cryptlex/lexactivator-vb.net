Imports System.IO
Imports Sample.Cryptlex

Public Class Form1
    Public Sub New()
        InitializeComponent()
        Dim status As Integer
        status = LexActivator.SetProductFile("Product.dat")
        If status <> LexActivator.LA_OK Then
            Me.statusLabel.Text = "Error setting product file: " + status.ToString()
            Return
        End If
        status = LexActivator.SetVersionGUID("59A44CE9-5415-8CF3-BD54-EA73A64E9A1B", LexActivator.PermissionFlags.LA_USER)
        If status <> LexActivator.LA_OK Then
            Me.statusLabel.Text = "Error setting version GUID: " + status.ToString()
            Return
        End If
        status = LexActivator.IsProductGenuine()
        If status = LexActivator.LA_OK OrElse status = LexActivator.LA_GP_OVER Then
            Me.statusLabel.Text = "Product genuinely activated!"
            Me.activateBtn.Text = "Deactivate"
            Me.activateTrialBtn.Enabled = False
            Return
        End If
        status = LexActivator.IsTrialGenuine()
        If status = LexActivator.LA_OK Then
            Dim daysLeft As UInteger = 0
            LexActivator.GetTrialDaysLeft(daysLeft, LexActivator.TrialType.LA_V_TRIAL)
            Me.statusLabel.Text = "Trial period! Days left:" + daysLeft.ToString()
            Me.activateTrialBtn.Enabled = False
        ElseIf status = LexActivator.LA_T_EXPIRED Then
            Me.statusLabel.Text = "Trial has expired!"
        Else
            Me.statusLabel.Text = "Trial has not started or has been tampered: " + status.ToString()
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
            Me.statusLabel.Text = "Error deactivating product: " + status.ToString()
            Return
        End If
        status = LexActivator.SetProductKey(productKeyBox.Text)
        If status <> LexActivator.LA_OK Then
            Me.statusLabel.Text = "Error setting product key: " + status.ToString()
            Return
        End If
        LexActivator.SetExtraActivationData("sample data")
        status = LexActivator.ActivateProduct()
        If status <> LexActivator.LA_OK Then
            Me.statusLabel.Text = "Error activating the product: " + status.ToString()
            Return
        Else
            Me.statusLabel.Text = "Activation Successful"
            Me.activateBtn.Text = "Deactivate"
            Me.activateTrialBtn.Enabled = False
        End If
    End Sub

    Public Sub activateTrialBtn_Click(sender As Object, e As EventArgs) Handles activateTrialBtn.Click
        Dim status As Integer
        status = LexActivator.SetTrialKey("CCEAF69B-144EDE48-B763AE2F-A0957C93-98827434")
        If status <> LexActivator.LA_OK Then
            Me.statusLabel.Text = "Error setting trial key: " + status.ToString()
            Return
        End If
        status = LexActivator.ActivateTrial()
        If status <> LexActivator.LA_OK Then
            Me.statusLabel.Text = "Error activating the trial: " + status.ToString()
            Return
        Else
            Me.statusLabel.Text = "Trial started Successful"
        End If
    End Sub
End Class
