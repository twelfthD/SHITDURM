Imports System.Drawing.Printing
Imports System.Data.SqlClient
Imports System.Printing


Public Class Form1
    Inherits System.Windows.Forms.Form

    ' Constant variable holding the Printer name.
    ' Private Const PRINTER_NAME As String = "Microsoft Print to PDF"
    Private Const PRINTER_NAME As String = "EPSON TM-T82 Receipt"
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents pbImage As PictureBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents txtVehicleno As TextBox
    Friend WithEvents TextBoxAmount As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents LabelSlno As Label
    Friend WithEvents txtDateTime As DateTimePicker
    Friend WithEvents btnExport As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents SHITDUDBDataSet As SHITDUDBDataSet
    Friend WithEvents TbRecieptBindingSource As BindingSource
    Friend WithEvents TbRecieptTableAdapter As SHITDUDBDataSetTableAdapters.tbRecieptTableAdapter
    Friend WithEvents IdDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents DateTimeDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents VehicleNoDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents AmountDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents btnRefresh As Button

    ' Variables/Objects.
    Private WithEvents pdPrint As PrintDocument

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdPrint As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.btnExport = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.IdDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DateTimeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.VehicleNoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AmountDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TbRecieptBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.SHITDUDBDataSet = New Program04.SHITDUDBDataSet()
        Me.txtDateTime = New System.Windows.Forms.DateTimePicker()
        Me.LabelSlno = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TextBoxAmount = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtVehicleno = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pbImage = New System.Windows.Forms.PictureBox()
        Me.cmdPrint = New System.Windows.Forms.Button()
        Me.TbRecieptTableAdapter = New Program04.SHITDUDBDataSetTableAdapters.tbRecieptTableAdapter()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TbRecieptBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SHITDUDBDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbImage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnRefresh)
        Me.GroupBox1.Controls.Add(Me.btnExport)
        Me.GroupBox1.Controls.Add(Me.DataGridView1)
        Me.GroupBox1.Controls.Add(Me.txtDateTime)
        Me.GroupBox1.Controls.Add(Me.LabelSlno)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.TextBoxAmount)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.txtVehicleno)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.pbImage)
        Me.GroupBox1.Controls.Add(Me.cmdPrint)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(756, 329)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Print Receipt"
        '
        'btnRefresh
        '
        Me.btnRefresh.Location = New System.Drawing.Point(586, 298)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(75, 23)
        Me.btnRefresh.TabIndex = 25
        Me.btnRefresh.Text = "Refresh"
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'btnExport
        '
        Me.btnExport.Location = New System.Drawing.Point(667, 298)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(75, 23)
        Me.btnExport.TabIndex = 24
        Me.btnExport.Text = "Export"
        Me.btnExport.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.IdDataGridViewTextBoxColumn, Me.DateTimeDataGridViewTextBoxColumn, Me.VehicleNoDataGridViewTextBoxColumn, Me.AmountDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.TbRecieptBindingSource
        Me.DataGridView1.Location = New System.Drawing.Point(300, 19)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(442, 273)
        Me.DataGridView1.TabIndex = 23
        '
        'IdDataGridViewTextBoxColumn
        '
        Me.IdDataGridViewTextBoxColumn.DataPropertyName = "Id"
        Me.IdDataGridViewTextBoxColumn.HeaderText = "Id"
        Me.IdDataGridViewTextBoxColumn.Name = "IdDataGridViewTextBoxColumn"
        Me.IdDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DateTimeDataGridViewTextBoxColumn
        '
        Me.DateTimeDataGridViewTextBoxColumn.DataPropertyName = "DateTime"
        Me.DateTimeDataGridViewTextBoxColumn.HeaderText = "DateTime"
        Me.DateTimeDataGridViewTextBoxColumn.Name = "DateTimeDataGridViewTextBoxColumn"
        '
        'VehicleNoDataGridViewTextBoxColumn
        '
        Me.VehicleNoDataGridViewTextBoxColumn.DataPropertyName = "VehicleNo"
        Me.VehicleNoDataGridViewTextBoxColumn.HeaderText = "VehicleNo"
        Me.VehicleNoDataGridViewTextBoxColumn.Name = "VehicleNoDataGridViewTextBoxColumn"
        '
        'AmountDataGridViewTextBoxColumn
        '
        Me.AmountDataGridViewTextBoxColumn.DataPropertyName = "Amount"
        Me.AmountDataGridViewTextBoxColumn.HeaderText = "Amount"
        Me.AmountDataGridViewTextBoxColumn.Name = "AmountDataGridViewTextBoxColumn"
        '
        'TbRecieptBindingSource
        '
        Me.TbRecieptBindingSource.DataMember = "tbReciept"
        Me.TbRecieptBindingSource.DataSource = Me.SHITDUDBDataSet
        '
        'SHITDUDBDataSet
        '
        Me.SHITDUDBDataSet.DataSetName = "SHITDUDBDataSet"
        Me.SHITDUDBDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'txtDateTime
        '
        Me.txtDateTime.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.txtDateTime.Location = New System.Drawing.Point(88, 106)
        Me.txtDateTime.Name = "txtDateTime"
        Me.txtDateTime.Size = New System.Drawing.Size(99, 20)
        Me.txtDateTime.TabIndex = 22
        '
        'LabelSlno
        '
        Me.LabelSlno.AutoSize = True
        Me.LabelSlno.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelSlno.Location = New System.Drawing.Point(97, 151)
        Me.LabelSlno.Name = "LabelSlno"
        Me.LabelSlno.Size = New System.Drawing.Size(0, 16)
        Me.LabelSlno.TabIndex = 21
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label7.Location = New System.Drawing.Point(315, 310)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(143, 13)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "© 2022 twelthd@github.com"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label6.Location = New System.Drawing.Point(294, 297)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(185, 13)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "SHITDU Receipt Maker - Ver. 1.1.0.3"
        '
        'TextBoxAmount
        '
        Me.TextBoxAmount.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxAmount.Location = New System.Drawing.Point(98, 209)
        Me.TextBoxAmount.Name = "TextBoxAmount"
        Me.TextBoxAmount.Size = New System.Drawing.Size(130, 20)
        Me.TextBoxAmount.TabIndex = 16
        Me.TextBoxAmount.Text = "50"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label5.Location = New System.Drawing.Point(62, 212)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(29, 13)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "Rs. :"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtVehicleno
        '
        Me.txtVehicleno.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVehicleno.Location = New System.Drawing.Point(98, 178)
        Me.txtVehicleno.Name = "txtVehicleno"
        Me.txtVehicleno.Size = New System.Drawing.Size(130, 20)
        Me.txtVehicleno.TabIndex = 14
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label4.Location = New System.Drawing.Point(23, 181)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(68, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Vehicle No. :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label3.Location = New System.Drawing.Point(45, 151)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(46, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "SL.No. :"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label2.Location = New System.Drawing.Point(28, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(218, 39)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Head Office : Kangpokpi" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Regd. No. : 1038 of 2022. (Act, XVI of 1926)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "AITUC, New" &
    " Delhi"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label1.Location = New System.Drawing.Point(19, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(236, 26)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "SADAR HILLS INLAND" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "TRANSPORTER AND DRIVER'S UNION" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pbImage
        '
        Me.pbImage.Image = CType(resources.GetObject("pbImage.Image"), System.Drawing.Image)
        Me.pbImage.Location = New System.Drawing.Point(97, 10)
        Me.pbImage.Name = "pbImage"
        Me.pbImage.Size = New System.Drawing.Size(80, 80)
        Me.pbImage.TabIndex = 4
        Me.pbImage.TabStop = False
        Me.pbImage.Visible = False
        '
        'cmdPrint
        '
        Me.cmdPrint.BackColor = System.Drawing.Color.GreenYellow
        Me.cmdPrint.Font = New System.Drawing.Font("Arial Rounded MT Bold", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPrint.Location = New System.Drawing.Point(97, 253)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(103, 57)
        Me.cmdPrint.TabIndex = 0
        Me.cmdPrint.Text = "Print"
        Me.cmdPrint.UseVisualStyleBackColor = False
        '
        'TbRecieptTableAdapter
        '
        Me.TbRecieptTableAdapter.ClearBeforeFill = True
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(771, 347)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SHITDU RECIEPT MAKER"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TbRecieptBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SHITDUDBDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbImage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim conn As New SqlConnection("Data Source=(localdb)\MSSQLLocalDB; AttachDbFilename=|DataDirectory|\SHITDUDB.MDF;Integrated Security=true")
    Public Sub ExecuteQuery(ByVal query As String)
        Dim cmd As New SqlCommand(query, conn)
        conn.Open()
        cmd.ExecuteNonQuery()
        conn.Close()
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load, btnRefresh.Click
        'TODO: This line of code loads data into the 'SHITDUDBDataSet.tbReciept' table. You can move, or remove it, as needed.
        Me.TbRecieptTableAdapter.Fill(Me.SHITDUDBDataSet.tbReciept)

        Dim loadSlno As String = "SELECT MAX(Id) FROM tbReciept"
        Dim cmd As New SqlCommand(loadSlno, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        da.Fill(dt)
        Dim getCount As String = dt.Rows(0)(0).ToString
        If getCount = "" Then
            getCount = "0"
        End If
        Dim counter As Int32 = CInt(getCount)

        counter = counter + 1
        LabelSlno.Text = counter.ToString

    End Sub
    ' The executed function when the Print button is clicked.
    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrint.Click
        pdPrint = New PrintDocument
        ' Change the printer to the indicated printer
        pdPrint.PrinterSettings.PrinterName = PRINTER_NAME

        Dim userPrinterName As New PrinterSettings()
        Dim myDefaultQueue As PrintQueue = Nothing

        Dim localPrintServer As New LocalPrintServer()
        ' Retrieving collection of local printer on user machine
        myDefaultQueue = LocalPrintServer.GetDefaultPrintQueue

        myDefaultQueue.Refresh()

        Dim uPrintername As String = "EPSON TM-T82 Receipt"
        If userPrinterName.PrinterName = uPrintername Then

            If myDefaultQueue.IsNotAvailable Then
                MessageBox.Show("Your printer is Offline or Not Available")
            Else

                'MessageBox.Show("Your printer is Online")
                If pdPrint.PrinterSettings.IsValid Then
                    'If myDefaultQueue.IsNotAvailable And myDefaultQueue.IsOffline = False Then

                    'MessageBox.Show("Your printer is offline")

                    'End If

                    If MessageBox.Show("Save Record and Print", "CONFIRM", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
                        pdPrint.DocumentName = "Printing SlNo. " + LabelSlno.Text.ToString

                        Dim insertquery As String = "INSERT INTO tbReciept(DateTime,VehicleNo,Amount)VALUES('" & txtDateTime.Value & "','" & txtVehicleno.Text & "','" & TextBoxAmount.Text & "')"
                        ExecuteQuery(insertquery)


                        Dim counter As String = LabelSlno.Text
                        Dim counter2int As Int32 = CInt(counter)
                        counter2int = counter2int
                        LabelSlno.Text = counter2int.ToString

                        ' Start printing
                        pdPrint.Print()

                        counter2int = counter2int + 1
                        LabelSlno.Text = counter2int.ToString
                        txtVehicleno.Clear()
                    End If
                End If
            End If
        Else
            MessageBox.Show("Please set 'EPSON TM-T82 Receipt' as your Default Printer. Your Default Printer is '" + userPrinterName.PrinterName.ToString + "'")
        End If



    End Sub

    ' The event handler function when pdPrint.Print is called.
    ' This is where the actual printing of sample data to the printer is made.
    Private Sub pdPrint_Print(ByVal sender As System.Object, ByVal e As PrintPageEventArgs) Handles pdPrint.PrintPage
        Dim x, y, lineOffset As Integer

        ' Instantiate font objects used in printing.
        Dim printFont As New Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point) 'Substituted to FontA Font
        Dim printFont2 As New Font("Microsoft Sans Serif", 16, FontStyle.Regular, GraphicsUnit.Point) 'Substituted to FontA Font

        e.Graphics.PageUnit = GraphicsUnit.Point



        ' Print the logo
        lineOffset = printFont.GetHeight(e.Graphics) - 2
        x = 10
        y = 5 + lineOffset

        ' Print the date and time
        Dim dt As DateTime = Now
        Dim fmt_dt As String = "        Date: " + dt
        e.Graphics.DrawString(fmt_dt, printFont, Brushes.Black, x, y)
        y += lineOffset
        y += lineOffset

        e.Graphics.DrawString("SL.N0.: " + LabelSlno.Text, printFont2, Brushes.Black, x, y)
        y += lineOffset + (lineOffset * 1.5)
        e.Graphics.DrawString("Vehicle No: " + txtVehicleno.Text, printFont2, Brushes.Black, x, y)
        y += lineOffset + (lineOffset * 1.5)
        e.Graphics.DrawString("Rs.: " + TextBoxAmount.Text + "/-", printFont2, Brushes.Black, x, y)
        y += lineOffset + (lineOffset * 2)
        e.Graphics.DrawString("Thanks for your Co-operation !", printFont, Brushes.Black, x, y)
        y += lineOffset
        e.Graphics.DrawString("**********************************", printFont2, Brushes.Black, x, y)
        y += lineOffset

        ' Indicate that no more data to print, and the Print Document can now send the print data to the spooler.
        e.HasMorePages = False



    End Sub

    ' The executed function when the Close button is clicked.
    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Close()
    End Sub




    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs)

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles txtDateTime.ValueChanged

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub pbImage_Click(sender As Object, e As EventArgs) Handles pbImage.Click

    End Sub


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox1_TextChanged_1(sender As Object, e As EventArgs) Handles TextBoxAmount.TextChanged

    End Sub


    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

    End Sub

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Private Sub BtnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Try
            Dim xlApp As Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
            Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim misValue As Object = System.Reflection.Missing.Value
            Dim i As Integer
            Dim j As Integer
            xlApp = New Microsoft.Office.Interop.Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheet = xlWorkBook.Sheets("sheet1")
            xlWorkSheet.Columns.AutoFit()
            For i = 0 To DataGridView1.RowCount - 2
                For j = 0 To DataGridView1.ColumnCount - 1
                    For k As Integer = 1 To DataGridView1.Columns.Count
                        xlWorkSheet.Cells(1, k) = DataGridView1.Columns(k - 1).HeaderText
                        xlWorkSheet.Cells(i + 2, j + 1) = DataGridView1(j, i).Value.ToString()
                    Next
                Next
            Next
            Dim fName As String = "DataBuku"
            Using sfd As New SaveFileDialog
                sfd.Title = "Save As"
                sfd.OverwritePrompt = True
                sfd.FileName = fName
                sfd.DefaultExt = ".xlsx"
                sfd.Filter = "Excel Workbook(*.xlsx)|"
                sfd.AddExtension = True
                If sfd.ShowDialog() = DialogResult.OK Then
                    xlWorkSheet.SaveAs(sfd.FileName)
                    xlWorkBook.Close()
                    xlApp.Quit()
                    releaseObject(xlApp)
                    releaseObject(xlWorkBook)
                    releaseObject(xlWorkSheet)
                    MsgBox("Database export success !", MsgBoxStyle.Information, "Export to Excel Worksheet")
                End If
            End Using
        Catch ex As Exception
            conn.Close()
            MsgBox("Export error ! " & vbCrLf & "Code error: " & ex.Message)
        End Try
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs)

    End Sub
End Class
