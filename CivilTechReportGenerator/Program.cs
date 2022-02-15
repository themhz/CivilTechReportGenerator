using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autofac;
using CivilTechReportGenerator.Handlers;
using CivilTechReportGenerator.Interfaces;
using CivilTechReportGenerator.Types;
using DevExpress.XtraRichEdit;

namespace CivilTechReportGenerator
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
                var app = scope.Resolve<IApp>();
                var listItem = scope.Resolve<IListItem>();
                var paragraphItem = scope.Resolve<IParagraphItem>();
                var sectionItem = scope.Resolve<ISectionItem>();
                var tableItem = scope.Resolve<ITableItem>();
                var listTableData = scope.Resolve<List<TableData>>();

                //List<TableData>

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
            builder.RegisterType<App>().As<IApp>();
            builder.RegisterType<RichEditDocumentServer>();
            builder.RegisterType<List<TableData>>();

            return builder.Build();
        }
    }
}
