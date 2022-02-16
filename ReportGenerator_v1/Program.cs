using System;
using System.Reflection;
using Autofac;
using DevExpress.XtraRichEdit;
using ReportGenerator_v1.System;

namespace ReportGenerator_v1 {
    class Program {
        public static IContainer Container { get; set; }
        static void Main(string[] args) {

            var Container = BuildContainer();
            using (var scope = Container.BeginLifetimeScope()) {
                var app = scope.Resolve<App>();
                app.start();
            }



        }

        private static IContainer BuildContainer() {
                var builder = new ContainerBuilder();
                builder.RegisterAssemblyTypes(Assembly.GetExecutingAssembly()).AsSelf().AsImplementedInterfaces();

                builder.RegisterType<RichEditDocumentServer>();
            

                return builder.Build();
        }
    }
}
