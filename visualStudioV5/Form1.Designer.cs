namespace smartFridge_v02
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.labelAskForItem = new System.Windows.Forms.Label();
            this.labelPutIn = new System.Windows.Forms.Label();
            this.tbAskForItem = new System.Windows.Forms.TextBox();
            this.tbPutIn = new System.Windows.Forms.TextBox();
            this.tbMessages = new System.Windows.Forms.TextBox();
            this.buttonAskForItem = new System.Windows.Forms.Button();
            this.buttonPutIn = new System.Windows.Forms.Button();
            this.serialPort1 = new System.IO.Ports.SerialPort(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.viewContents = new System.Windows.Forms.Button();
            this.bShowCal = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // labelAskForItem
            // 
            this.labelAskForItem.AutoSize = true;
            this.labelAskForItem.Location = new System.Drawing.Point(42, 57);
            this.labelAskForItem.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelAskForItem.Name = "labelAskForItem";
            this.labelAskForItem.Size = new System.Drawing.Size(262, 20);
            this.labelAskForItem.TabIndex = 0;
            this.labelAskForItem.Text = "Enter an Item to get from the fridge:";
            // 
            // labelPutIn
            // 
            this.labelPutIn.AutoSize = true;
            this.labelPutIn.Location = new System.Drawing.Point(429, 57);
            this.labelPutIn.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelPutIn.Name = "labelPutIn";
            this.labelPutIn.Size = new System.Drawing.Size(256, 20);
            this.labelPutIn.TabIndex = 1;
            this.labelPutIn.Text = "Enter an Item to put into the fridge:";
            // 
            // tbAskForItem
            // 
            this.tbAskForItem.Location = new System.Drawing.Point(40, 109);
            this.tbAskForItem.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.tbAskForItem.MaxLength = 48;
            this.tbAskForItem.Multiline = true;
            this.tbAskForItem.Name = "tbAskForItem";
            this.tbAskForItem.Size = new System.Drawing.Size(298, 95);
            this.tbAskForItem.TabIndex = 2;
            // 
            // tbPutIn
            // 
            this.tbPutIn.Location = new System.Drawing.Point(431, 110);
            this.tbPutIn.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.tbPutIn.MaxLength = 48;
            this.tbPutIn.Multiline = true;
            this.tbPutIn.Name = "tbPutIn";
            this.tbPutIn.Size = new System.Drawing.Size(298, 94);
            this.tbPutIn.TabIndex = 3;
            // 
            // tbMessages
            // 
            this.tbMessages.Location = new System.Drawing.Point(40, 325);
            this.tbMessages.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.tbMessages.Multiline = true;
            this.tbMessages.Name = "tbMessages";
            this.tbMessages.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tbMessages.Size = new System.Drawing.Size(589, 156);
            this.tbMessages.TabIndex = 4;
            // 
            // buttonAskForItem
            // 
            this.buttonAskForItem.Location = new System.Drawing.Point(40, 223);
            this.buttonAskForItem.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.buttonAskForItem.Name = "buttonAskForItem";
            this.buttonAskForItem.Size = new System.Drawing.Size(92, 41);
            this.buttonAskForItem.TabIndex = 5;
            this.buttonAskForItem.Text = "Submit";
            this.buttonAskForItem.UseVisualStyleBackColor = true;
            this.buttonAskForItem.Click += new System.EventHandler(this.buttonAskForItem_Click);
            // 
            // buttonPutIn
            // 
            this.buttonPutIn.Location = new System.Drawing.Point(431, 223);
            this.buttonPutIn.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.buttonPutIn.Name = "buttonPutIn";
            this.buttonPutIn.Size = new System.Drawing.Size(92, 41);
            this.buttonPutIn.TabIndex = 6;
            this.buttonPutIn.Text = "Submit";
            this.buttonPutIn.UseVisualStyleBackColor = true;
            this.buttonPutIn.Click += new System.EventHandler(this.buttonPutIn_Click);
            // 
            // serialPort1
            // 
            this.serialPort1.BaudRate = 115200;
            this.serialPort1.PortName = "COM3";
            this.serialPort1.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(this.serialPort1_DataReceived);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(61, 343);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 20);
            this.label1.TabIndex = 7;
            // 
            // viewContents
            // 
            this.viewContents.Location = new System.Drawing.Point(431, 325);
            this.viewContents.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.viewContents.Name = "viewContents";
            this.viewContents.Size = new System.Drawing.Size(240, 46);
            this.viewContents.TabIndex = 8;
            this.viewContents.Text = "View contents of fridge";
            this.viewContents.UseVisualStyleBackColor = true;
            this.viewContents.Click += new System.EventHandler(this.viewContents_Click);
            // 
            // bShowCal
            // 
            this.bShowCal.Location = new System.Drawing.Point(1097, 471);
            this.bShowCal.Name = "bShowCal";
            this.bShowCal.Size = new System.Drawing.Size(177, 67);
            this.bShowCal.TabIndex = 11;
            this.bShowCal.Text = "Show the Calendar";
            this.bShowCal.UseVisualStyleBackColor = true;
            this.bShowCal.Click += new System.EventHandler(this.bShowCal_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1364, 715);
            this.Controls.Add(this.bShowCal);
            this.Controls.Add(this.viewContents);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonPutIn);
            this.Controls.Add(this.buttonAskForItem);
            this.Controls.Add(this.tbMessages);
            this.Controls.Add(this.tbPutIn);
            this.Controls.Add(this.tbAskForItem);
            this.Controls.Add(this.labelPutIn);
            this.Controls.Add(this.labelAskForItem);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "Form1";
            this.Text = "The Smartest Fridge";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelAskForItem;
        private System.Windows.Forms.Label labelPutIn;
        private System.Windows.Forms.TextBox tbAskForItem;
        private System.Windows.Forms.TextBox tbPutIn;
        private System.Windows.Forms.TextBox tbMessages;
        private System.Windows.Forms.Button buttonAskForItem;
        private System.Windows.Forms.Button buttonPutIn;
        private System.IO.Ports.SerialPort serialPort1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button viewContents;
        private System.Windows.Forms.Button bShowCal;
    }
}

