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
    Friend WithEvents PictureBox1 As PictureBox

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

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdPrint As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
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
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.GroupBox1.SuspendLayout()
        CType(Me.pbImage, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.Color.Tomato
        Me.cmdClose.Location = New System.Drawing.Point(134, 256)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(120, 65)
        Me.cmdClose.TabIndex = 3
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.PictureBox1)
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
        Me.GroupBox1.Controls.Add(Me.cmdClose)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(559, 329)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Print Receipt"
        '
        'txtDateTime
        '
        Me.txtDateTime.Location = New System.Drawing.Point(69, 106)
        Me.txtDateTime.Name = "txtDateTime"
        Me.txtDateTime.Size = New System.Drawing.Size(129, 20)
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
        Me.Label7.Location = New System.Drawing.Point(339, 308)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(143, 13)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "© 2022 twelthd@github.com"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label6.Location = New System.Drawing.Point(322, 295)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(176, 13)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "SHITDU Receipt Maker - Ver. 1.0.0.6"
        '
        'TextBoxAmount
        '
        Me.TextBoxAmount.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxAmount.Location = New System.Drawing.Point(98, 209)
        Me.TextBoxAmount.Name = "TextBoxAmount"
        Me.TextBoxAmount.Size = New System.Drawing.Size(130, 20)
        Me.TextBoxAmount.TabIndex = 16
        Me.TextBoxAmount.Text = "100"
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
        Me.Label2.Location = New System.Drawing.Point(24, 51)
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
        Me.Label1.Size = New System.Drawing.Size(228, 26)
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
        Me.cmdPrint.Location = New System.Drawing.Point(13, 256)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.Size = New System.Drawing.Size(120, 65)
        Me.cmdPrint.TabIndex = 0
        Me.cmdPrint.Text = "Print"
        Me.cmdPrint.UseVisualStyleBackColor = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.Program04.My.Resources.Resource1.shidu_logo
        Me.PictureBox1.Location = New System.Drawing.Point(287, 19)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(257, 258)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 3
        Me.PictureBox1.TabStop = False
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(573, 347)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SHITDU RECIEPT MAKER"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.pbImage, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
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
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
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


        Dim myDefaultQueue As PrintQueue = Nothing

        Dim localPrintServer As New LocalPrintServer()
        ' Retrieving collection of local printer on user machine
        myDefaultQueue = LocalPrintServer.GetDefaultPrintQueue

        myDefaultQueue.Refresh()
        If myDefaultQueue.IsNotAvailable And myDefaultQueue.IsOffline = False Then
            MessageBox.Show("Your printer is offline")
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


            'MessageBox.Show("Print Cancled")
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
    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
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
End Class
