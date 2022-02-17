using Autofac;
using DevExpress.XtraRichEdit;
using ReportGenerator_v1.System;
using System.Reflection;

namespace ReportGenerator_v1 {
    class Program {
        public static IContainer Container { get; set; }
        static void Main(string[] args) {

            var Container = BuildContainer();
            using (var scope = Container.BeginLifetimeScope()) {
                var app = scope.Resolve<App>();
                //var reportType = scope.Resolve<ExceedDocX>();
                var reportType = scope.Resolve<DevExpressDocX>();

                app.start(reportType);
            }
        }

        //This is the container for the autofac. You register all your objects here
        //https://autofac.readthedocs.io/en/latest/index.html
        private static IContainer BuildContainer() {
            var builder = new ContainerBuilder();
            builder.RegisterAssemblyTypes(Assembly.GetExecutingAssembly()).AsSelf().AsImplementedInterfaces();
            builder.RegisterType<RichEditDocumentServer>();
            builder.RegisterType<ExceedDocX>();

            return builder.Build();
        }
    }
}
