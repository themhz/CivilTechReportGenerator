using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autofac;
using CivilTechReportGenerator.Handlers;
using CivilTechReportGenerator.Interfaces;

namespace CivilTechReportGenerator
{
    static class Program
    {
        private static IContainer Container { get; set; }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            var builder = new ContainerBuilder();
            builder.RegisterType<ListItem>().As<IListItem>();
            builder.RegisterType<ParagraphItem>().As<IParagraphItem>();
            builder.RegisterType<SectionItem>().As<ISectionItem>();
            builder.RegisterType<TableItem>().As<ITableItem>();

            //builder.RegisterType<DocumentX>().As<IDocumentX>();
            //builder.RegisterType<DocumentXItem>().As<IDocumentXItem>();

            Container = builder.Build();


            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
