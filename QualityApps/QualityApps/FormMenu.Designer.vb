<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormMenu
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormMenu))
        Me.ToolStripContainer1 = New System.Windows.Forms.ToolStripContainer()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.TransactionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportSQ01ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GenerateExcelForSupplierToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GenerateExcelForSupplierToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.CollectReplyFromSupplierToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Send2DaysEmailInAdvanceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.InspectionReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ActivityLogToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CreateActivityLogToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportActivityLogToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportAcitivtyLogAllDataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MasterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.VendorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AnnouncementToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ParameterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UserToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ActivityToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.VendorAssignmentQEUserToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FirstCmmfToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MissingVendorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AdminToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RBACToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OTMToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportDataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GenerateCSVToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.InspectionReportToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripContainer1.BottomToolStripPanel.SuspendLayout()
        Me.ToolStripContainer1.TopToolStripPanel.SuspendLayout()
        Me.ToolStripContainer1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStripContainer1
        '
        '
        'ToolStripContainer1.BottomToolStripPanel
        '
        Me.ToolStripContainer1.BottomToolStripPanel.Controls.Add(Me.StatusStrip1)
        '
        'ToolStripContainer1.ContentPanel
        '
        Me.ToolStripContainer1.ContentPanel.Size = New System.Drawing.Size(556, 80)
        Me.ToolStripContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStripContainer1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStripContainer1.Name = "ToolStripContainer1"
        Me.ToolStripContainer1.Size = New System.Drawing.Size(556, 126)
        Me.ToolStripContainer1.TabIndex = 0
        Me.ToolStripContainer1.Text = "ToolStripContainer1"
        '
        'ToolStripContainer1.TopToolStripPanel
        '
        Me.ToolStripContainer1.TopToolStripPanel.Controls.Add(Me.MenuStrip1)
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Dock = System.Windows.Forms.DockStyle.None
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(556, 22)
        Me.StatusStrip1.TabIndex = 0
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Dock = System.Windows.Forms.DockStyle.None
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TransactionToolStripMenuItem, Me.ReportsToolStripMenuItem, Me.ActivityLogToolStripMenuItem, Me.MasterToolStripMenuItem, Me.AdminToolStripMenuItem, Me.OTMToolStripMenuItem, Me.ReportToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(556, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'TransactionToolStripMenuItem
        '
        Me.TransactionToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ImportSQ01ToolStripMenuItem, Me.GenerateExcelForSupplierToolStripMenuItem, Me.GenerateExcelForSupplierToolStripMenuItem1, Me.CollectReplyFromSupplierToolStripMenuItem, Me.Send2DaysEmailInAdvanceToolStripMenuItem})
        Me.TransactionToolStripMenuItem.Name = "TransactionToolStripMenuItem"
        Me.TransactionToolStripMenuItem.Size = New System.Drawing.Size(59, 20)
        Me.TransactionToolStripMenuItem.Text = "Actions"
        '
        'ImportSQ01ToolStripMenuItem
        '
        Me.ImportSQ01ToolStripMenuItem.Name = "ImportSQ01ToolStripMenuItem"
        Me.ImportSQ01ToolStripMenuItem.Size = New System.Drawing.Size(580, 22)
        Me.ImportSQ01ToolStripMenuItem.Tag = "FormImportDailyExtraction"
        Me.ImportSQ01ToolStripMenuItem.Text = "Import SQ01-F037-Quality-002-PO Confirmation-Inspection Lot (Local File - Spreads" & _
    "heet - 1133)"
        '
        'GenerateExcelForSupplierToolStripMenuItem
        '
        Me.GenerateExcelForSupplierToolStripMenuItem.Name = "GenerateExcelForSupplierToolStripMenuItem"
        Me.GenerateExcelForSupplierToolStripMenuItem.Size = New System.Drawing.Size(580, 22)
        Me.GenerateExcelForSupplierToolStripMenuItem.Tag = "FormGenerateExcel"
        Me.GenerateExcelForSupplierToolStripMenuItem.Text = "Generate Excel For Supplier"
        Me.GenerateExcelForSupplierToolStripMenuItem.Visible = False
        '
        'GenerateExcelForSupplierToolStripMenuItem1
        '
        Me.GenerateExcelForSupplierToolStripMenuItem1.Name = "GenerateExcelForSupplierToolStripMenuItem1"
        Me.GenerateExcelForSupplierToolStripMenuItem1.Size = New System.Drawing.Size(580, 22)
        Me.GenerateExcelForSupplierToolStripMenuItem1.Tag = "FormGenerateExcelDataGrid"
        Me.GenerateExcelForSupplierToolStripMenuItem1.Text = "Generate Excel For Supplier"
        '
        'CollectReplyFromSupplierToolStripMenuItem
        '
        Me.CollectReplyFromSupplierToolStripMenuItem.Name = "CollectReplyFromSupplierToolStripMenuItem"
        Me.CollectReplyFromSupplierToolStripMenuItem.Size = New System.Drawing.Size(580, 22)
        Me.CollectReplyFromSupplierToolStripMenuItem.Tag = "FormGetReply"
        Me.CollectReplyFromSupplierToolStripMenuItem.Text = "Collect Reply From Supplier"
        '
        'Send2DaysEmailInAdvanceToolStripMenuItem
        '
        Me.Send2DaysEmailInAdvanceToolStripMenuItem.Name = "Send2DaysEmailInAdvanceToolStripMenuItem"
        Me.Send2DaysEmailInAdvanceToolStripMenuItem.Size = New System.Drawing.Size(580, 22)
        Me.Send2DaysEmailInAdvanceToolStripMenuItem.Tag = "FormSendEmailConfirmation"
        Me.Send2DaysEmailInAdvanceToolStripMenuItem.Text = "Send Email Inspection Schedule Confirmation"
        '
        'ReportsToolStripMenuItem
        '
        Me.ReportsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.InspectionReportToolStripMenuItem})
        Me.ReportsToolStripMenuItem.Name = "ReportsToolStripMenuItem"
        Me.ReportsToolStripMenuItem.Size = New System.Drawing.Size(78, 20)
        Me.ReportsToolStripMenuItem.Text = "Scheduling"
        '
        'InspectionReportToolStripMenuItem
        '
        Me.InspectionReportToolStripMenuItem.Name = "InspectionReportToolStripMenuItem"
        Me.InspectionReportToolStripMenuItem.Size = New System.Drawing.Size(186, 22)
        Me.InspectionReportToolStripMenuItem.Tag = "FormInspection"
        Me.InspectionReportToolStripMenuItem.Text = "Inspection Allocation"
        '
        'ActivityLogToolStripMenuItem
        '
        Me.ActivityLogToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CreateActivityLogToolStripMenuItem, Me.ReportActivityLogToolStripMenuItem, Me.ReportAcitivtyLogAllDataToolStripMenuItem})
        Me.ActivityLogToolStripMenuItem.Name = "ActivityLogToolStripMenuItem"
        Me.ActivityLogToolStripMenuItem.Size = New System.Drawing.Size(82, 20)
        Me.ActivityLogToolStripMenuItem.Text = "Activity Log"
        '
        'CreateActivityLogToolStripMenuItem
        '
        Me.CreateActivityLogToolStripMenuItem.Name = "CreateActivityLogToolStripMenuItem"
        Me.CreateActivityLogToolStripMenuItem.Size = New System.Drawing.Size(219, 22)
        Me.CreateActivityLogToolStripMenuItem.Tag = "FormActivityLog"
        Me.CreateActivityLogToolStripMenuItem.Text = "Create Activity Log"
        '
        'ReportActivityLogToolStripMenuItem
        '
        Me.ReportActivityLogToolStripMenuItem.Name = "ReportActivityLogToolStripMenuItem"
        Me.ReportActivityLogToolStripMenuItem.Size = New System.Drawing.Size(219, 22)
        Me.ReportActivityLogToolStripMenuItem.Tag = "FormGenerateReportActivityLog"
        Me.ReportActivityLogToolStripMenuItem.Text = "Report Activity Log"
        '
        'ReportAcitivtyLogAllDataToolStripMenuItem
        '
        Me.ReportAcitivtyLogAllDataToolStripMenuItem.Name = "ReportAcitivtyLogAllDataToolStripMenuItem"
        Me.ReportAcitivtyLogAllDataToolStripMenuItem.Size = New System.Drawing.Size(219, 22)
        Me.ReportAcitivtyLogAllDataToolStripMenuItem.Text = "Report Acitivty Log All Data"
        '
        'MasterToolStripMenuItem
        '
        Me.MasterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.VendorToolStripMenuItem, Me.AnnouncementToolStripMenuItem, Me.ParameterToolStripMenuItem, Me.UserToolStripMenuItem, Me.ActivityToolStripMenuItem, Me.VendorAssignmentQEUserToolStripMenuItem, Me.FirstCmmfToolStripMenuItem, Me.MissingVendorToolStripMenuItem})
        Me.MasterToolStripMenuItem.Name = "MasterToolStripMenuItem"
        Me.MasterToolStripMenuItem.Size = New System.Drawing.Size(55, 20)
        Me.MasterToolStripMenuItem.Text = "Master"
        '
        'VendorToolStripMenuItem
        '
        Me.VendorToolStripMenuItem.Name = "VendorToolStripMenuItem"
        Me.VendorToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.VendorToolStripMenuItem.Tag = "FormVendor"
        Me.VendorToolStripMenuItem.Text = "Vendor"
        '
        'AnnouncementToolStripMenuItem
        '
        Me.AnnouncementToolStripMenuItem.Name = "AnnouncementToolStripMenuItem"
        Me.AnnouncementToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.AnnouncementToolStripMenuItem.Tag = "FormAnnouncement"
        Me.AnnouncementToolStripMenuItem.Text = "Announcement"
        '
        'ParameterToolStripMenuItem
        '
        Me.ParameterToolStripMenuItem.Name = "ParameterToolStripMenuItem"
        Me.ParameterToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.ParameterToolStripMenuItem.Tag = "FormParameters"
        Me.ParameterToolStripMenuItem.Text = "Parameter"
        '
        'UserToolStripMenuItem
        '
        Me.UserToolStripMenuItem.Name = "UserToolStripMenuItem"
        Me.UserToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.UserToolStripMenuItem.Tag = "FormUser"
        Me.UserToolStripMenuItem.Text = "User"
        '
        'ActivityToolStripMenuItem
        '
        Me.ActivityToolStripMenuItem.Name = "ActivityToolStripMenuItem"
        Me.ActivityToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.ActivityToolStripMenuItem.Tag = "FormActivityMaster"
        Me.ActivityToolStripMenuItem.Text = "Activity"
        '
        'VendorAssignmentQEUserToolStripMenuItem
        '
        Me.VendorAssignmentQEUserToolStripMenuItem.Name = "VendorAssignmentQEUserToolStripMenuItem"
        Me.VendorAssignmentQEUserToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.VendorAssignmentQEUserToolStripMenuItem.Tag = "FormVendorQEAssignment"
        Me.VendorAssignmentQEUserToolStripMenuItem.Text = "Vendor Assignment QE User"
        '
        'FirstCmmfToolStripMenuItem
        '
        Me.FirstCmmfToolStripMenuItem.Name = "FirstCmmfToolStripMenuItem"
        Me.FirstCmmfToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.FirstCmmfToolStripMenuItem.Tag = "FormFirstCMMF"
        Me.FirstCmmfToolStripMenuItem.Text = "First CMMF"
        '
        'MissingVendorToolStripMenuItem
        '
        Me.MissingVendorToolStripMenuItem.Name = "MissingVendorToolStripMenuItem"
        Me.MissingVendorToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.MissingVendorToolStripMenuItem.Tag = "FormMissingVendor"
        Me.MissingVendorToolStripMenuItem.Text = "Missing Vendor"
        '
        'AdminToolStripMenuItem
        '
        Me.AdminToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RBACToolStripMenuItem})
        Me.AdminToolStripMenuItem.Name = "AdminToolStripMenuItem"
        Me.AdminToolStripMenuItem.Size = New System.Drawing.Size(55, 20)
        Me.AdminToolStripMenuItem.Text = "Admin"
        '
        'RBACToolStripMenuItem
        '
        Me.RBACToolStripMenuItem.Name = "RBACToolStripMenuItem"
        Me.RBACToolStripMenuItem.Size = New System.Drawing.Size(104, 22)
        Me.RBACToolStripMenuItem.Text = "RBAC"
        '
        'OTMToolStripMenuItem
        '
        Me.OTMToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ImportDataToolStripMenuItem, Me.GenerateCSVToolStripMenuItem})
        Me.OTMToolStripMenuItem.Name = "OTMToolStripMenuItem"
        Me.OTMToolStripMenuItem.Size = New System.Drawing.Size(45, 20)
        Me.OTMToolStripMenuItem.Text = "OTM"
        '
        'ImportDataToolStripMenuItem
        '
        Me.ImportDataToolStripMenuItem.Name = "ImportDataToolStripMenuItem"
        Me.ImportDataToolStripMenuItem.Size = New System.Drawing.Size(145, 22)
        Me.ImportDataToolStripMenuItem.Tag = "FormImportOTM"
        Me.ImportDataToolStripMenuItem.Text = "Import Data"
        '
        'GenerateCSVToolStripMenuItem
        '
        Me.GenerateCSVToolStripMenuItem.Name = "GenerateCSVToolStripMenuItem"
        Me.GenerateCSVToolStripMenuItem.Size = New System.Drawing.Size(145, 22)
        Me.GenerateCSVToolStripMenuItem.Tag = "FormGenerateCSVOTM"
        Me.GenerateCSVToolStripMenuItem.Text = "Generate CSV"
        '
        'ReportToolStripMenuItem
        '
        Me.ReportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.InspectionReportToolStripMenuItem1})
        Me.ReportToolStripMenuItem.Name = "ReportToolStripMenuItem"
        Me.ReportToolStripMenuItem.Size = New System.Drawing.Size(54, 20)
        Me.ReportToolStripMenuItem.Text = "Report"
        '
        'InspectionReportToolStripMenuItem1
        '
        Me.InspectionReportToolStripMenuItem1.Name = "InspectionReportToolStripMenuItem1"
        Me.InspectionReportToolStripMenuItem1.Size = New System.Drawing.Size(167, 22)
        Me.InspectionReportToolStripMenuItem1.Text = "Inspection Report"
        '
        'FormMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(556, 126)
        Me.Controls.Add(Me.ToolStripContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "FormMenu"
        Me.Text = "FormMenu"
        Me.ToolStripContainer1.BottomToolStripPanel.ResumeLayout(False)
        Me.ToolStripContainer1.BottomToolStripPanel.PerformLayout()
        Me.ToolStripContainer1.TopToolStripPanel.ResumeLayout(False)
        Me.ToolStripContainer1.TopToolStripPanel.PerformLayout()
        Me.ToolStripContainer1.ResumeLayout(False)
        Me.ToolStripContainer1.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ToolStripContainer1 As System.Windows.Forms.ToolStripContainer
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents TransactionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportSQ01ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GenerateExcelForSupplierToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GenerateExcelForSupplierToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CollectReplyFromSupplierToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReportsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents InspectionReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MasterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents VendorToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AnnouncementToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ParameterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Send2DaysEmailInAdvanceToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AdminToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RBACToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UserToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ActivityLogToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CreateActivityLogToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReportActivityLogToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OTMToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportDataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GenerateCSVToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ActivityToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents VendorAssignmentQEUserToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReportAcitivtyLogAllDataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FirstCmmfToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MissingVendorToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents InspectionReportToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem

End Class
