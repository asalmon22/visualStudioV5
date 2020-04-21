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
            this.continueButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // labelAskForItem
            // 
            this.labelAskForItem.AutoSize = true;
            this.labelAskForItem.Location = new System.Drawing.Point(65, 83);
            this.labelAskForItem.Name = "labelAskForItem";
            this.labelAskForItem.Size = new System.Drawing.Size(387, 29);
            this.labelAskForItem.TabIndex = 0;
            this.labelAskForItem.Text = "Enter an Item to get from the fridge:";
            // 
            // labelPutIn
            // 
            this.labelPutIn.AutoSize = true;
            this.labelPutIn.Location = new System.Drawing.Point(667, 83);
            this.labelPutIn.Name = "labelPutIn";
            this.labelPutIn.Size = new System.Drawing.Size(377, 29);
            this.labelPutIn.TabIndex = 1;
            this.labelPutIn.Text = "Enter an Item to put into the fridge:";
            // 
            // tbAskForItem
            // 
            this.tbAskForItem.Location = new System.Drawing.Point(62, 158);
            this.tbAskForItem.MaxLength = 48;
            this.tbAskForItem.Multiline = true;
            this.tbAskForItem.Name = "tbAskForItem";
            this.tbAskForItem.Size = new System.Drawing.Size(461, 136);
            this.tbAskForItem.TabIndex = 2;
            // 
            // tbPutIn
            // 
            this.tbPutIn.Location = new System.Drawing.Point(670, 160);
            this.tbPutIn.MaxLength = 48;
            this.tbPutIn.Multiline = true;
            this.tbPutIn.Name = "tbPutIn";
            this.tbPutIn.Size = new System.Drawing.Size(461, 135);
            this.tbPutIn.TabIndex = 3;
            // 
            // tbMessages
            // 
            this.tbMessages.Location = new System.Drawing.Point(62, 471);
            this.tbMessages.Multiline = true;
            this.tbMessages.Name = "tbMessages";
            this.tbMessages.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tbMessages.Size = new System.Drawing.Size(914, 224);
            this.tbMessages.TabIndex = 4;
            // 
            // buttonAskForItem
            // 
            this.buttonAskForItem.Location = new System.Drawing.Point(62, 323);
            this.buttonAskForItem.Name = "buttonAskForItem";
            this.buttonAskForItem.Size = new System.Drawing.Size(143, 59);
            this.buttonAskForItem.TabIndex = 5;
            this.buttonAskForItem.Text = "Submit";
            this.buttonAskForItem.UseVisualStyleBackColor = true;
            this.buttonAskForItem.Click += new System.EventHandler(this.buttonAskForItem_Click);
            // 
            // buttonPutIn
            // 
            this.buttonPutIn.Location = new System.Drawing.Point(670, 323);
            this.buttonPutIn.Name = "buttonPutIn";
            this.buttonPutIn.Size = new System.Drawing.Size(143, 59);
            this.buttonPutIn.TabIndex = 6;
            this.buttonPutIn.Text = "Submit";
            this.buttonPutIn.UseVisualStyleBackColor = true;
            this.buttonPutIn.Click += new System.EventHandler(this.buttonPutIn_Click);
            // 
            // serialPort1
            // 
            this.serialPort1.PortName = "COM3";
            this.serialPort1.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(this.serialPort1_DataReceived);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(95, 497);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 29);
            this.label1.TabIndex = 7;
            // 
            // viewContents
            // 
            this.viewContents.Location = new System.Drawing.Point(1008, 471);
            this.viewContents.Name = "viewContents";
            this.viewContents.Size = new System.Drawing.Size(373, 67);
            this.viewContents.TabIndex = 8;
            this.viewContents.Text = "View contents of fridge";
            this.viewContents.UseVisualStyleBackColor = true;
            this.viewContents.Click += new System.EventHandler(this.viewContents_Click);
            // 
            // bShowCal
            // 
            this.bShowCal.Location = new System.Drawing.Point(1706, 683);
            this.bShowCal.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.bShowCal.Name = "bShowCal";
            this.bShowCal.Size = new System.Drawing.Size(275, 97);
            this.bShowCal.TabIndex = 11;
            this.bShowCal.Text = "Show the Calendar";
            this.bShowCal.UseVisualStyleBackColor = true;
            this.bShowCal.Click += new System.EventHandler(this.bShowCal_Click);
            // 
            // continueButton
            // 
            this.continueButton.Location = new System.Drawing.Point(1221, 177);
            this.continueButton.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.continueButton.Name = "continueButton";
            this.continueButton.Size = new System.Drawing.Size(213, 91);
            this.continueButton.TabIndex = 12;
            this.continueButton.Text = "continue";
            this.continueButton.UseVisualStyleBackColor = true;
            this.continueButton.Click += new System.EventHandler(this.continueButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 29F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1540, 829);
            this.Controls.Add(this.continueButton);
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
        private System.Windows.Forms.Button continueButton;
    }
}

