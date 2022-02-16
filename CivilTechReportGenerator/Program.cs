using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autofac;
using ReportGenerator.Handlers;
using ReportGenerator.Interfaces;
using ReportGenerator.Types;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace ReportGenerator
{
    static class Program
    {
        public static IContainer Container { get; set; }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            
            //builder.RegisterType<DocumentX>().As<IDocumentX>();
            //builder.RegisterType<DocumentXItem>().As<IDocumentXItem>();

            Container = BuildContainer();

            using (var scope = Container.BeginLifetimeScope()) {
                //scope.Resolve<Application>().Run();                
                var app = scope.Resolve<ITests>();             
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1(app));
            }

            
        }

        private static IContainer BuildContainer() {
            var builder = new ContainerBuilder();
            builder.RegisterAssemblyTypes(Assembly.GetExecutingAssembly()).AsSelf().AsImplementedInterfaces();

            builder.RegisterType<ListItem>().As<IListItem>();
            builder.RegisterType<ParagraphItem>().As<IParagraphItem>();
            builder.RegisterType<SectionItem>().As<ISectionItem>();
            builder.RegisterType<TableItem>().As<ITableItem>();
            builder.RegisterType<Tests>().As<ITests>();
            builder.RegisterType<RichEditDocumentServer>();
            builder.RegisterType<List<TableData>>();
            builder.RegisterType<Regex>();
            builder.RegisterType<DocumentHandler>().As<IDocumentHandler>();
            builder.RegisterType<TableData>();


            return builder.Build();
        }
    }
}
