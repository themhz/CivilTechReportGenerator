using ReportGenerator.Handlers;
using ReportGenerator.tests;
using ReportGenerator.Types;
using DevExpress.XtraRichEdit;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace ReportGenerator {
    public partial class Form1 : Form {
        int a = 0;
        public ITests app { get; }
        public Form1(ITests _app) {
            InitializeComponent();
            this.app = _app;
        }        

        private void button1_Click(object sender, EventArgs e) {
            app.run();
        }

        
    }
}
