
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
//using System.Drawing.Printing;
//using System.Drawing.Drawing2D;
//using System.Reflection;
//using System.Drawing.Imaging;
using CDIntfEx;
//using Microsoft.Win32;
//using System.IO;
//using System.Runtime.InteropServices;
//using System.Threading;
//using System.Text;


namespace PDFConverter
{
    /// <summary>
    /// Summary description for Form1.
    /// </summary>
    public class PDFConverter : System.Windows.Forms.Form
    {
        private System.ComponentModel.Container components = null;
        private System.Windows.Forms.Button btnBatchConvert;
        public PDFConverter()
        {
            InitializeComponent();
        }

     protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
     
        private void InitializeComponent()
        {
            this.btnBatchConvert = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnBatchConvert
            // 
            this.btnBatchConvert.Location = new System.Drawing.Point(286, 194);
            this.btnBatchConvert.Name = "btnBatchConvert";
            this.btnBatchConvert.Size = new System.Drawing.Size(144, 48);
            this.btnBatchConvert.TabIndex = 4;
            this.btnBatchConvert.Text = "Batch Convert";
            this.btnBatchConvert.Click += new System.EventHandler(this.btnBatchConvert_Click);
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(462, 267);
            this.Controls.Add(this.btnBatchConvert);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.Text = "Amyuni PDF Converter Sample";
            this.ResumeLayout(false);

        }
        #endregion

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.Run(new PDFConverter());
        }


        private void btnBatchConvert_Click(object sender, System.EventArgs e)
        {
            Amyuni amyuni = new Amyuni();
            amyuni.BatchConvert();

        }

    }
}
