The root of the application is the ReportGenerator. ReportGenerator_tests are the tests.
Program starts from program.cs
in the main we add the type of the docx reporting type generator that we will use.

Declaring this means that we use the DevExpress tools in order to parse and create DocX files
var reportType = scope.Resolve<DevExpressDocX>();

 The system folder contains the main application files
 App.cs starts the ReportGenerator.cs and creates a DocX report

 public void start(IReport reportType) {
            
            this.reportGenerator.CreateDocX(reportType);         <-- this will create a report
 }

 The parameter reportType is an interface that is passed(ibjected) from the program.cs to the app.cs


//From programm.cs   
var reportType = scope.Resolve<DevExpressDocX>();  
app.start(reportType);  

//To App.cs  
public void start(IReport reportType) {  
        this.reportGenerator.CreateDocX(reportType);              
}  

ReportGeerator is called after the start by the CreateDocX message  
 public void CreateDocX(IReport reportType) {  
            this.DocXReport = reportType;  
            this.DocXReport.template = "Some Path for the template";  
            this.DocXReport.generatedfile = "Some Path for the generated file of the report";  
            this.DocXReport = this.DocXReport.create();  
}  
  
This will cause the DocXReport to create the report depending on the parameter from the start of the application.  
In folder IReports we can add more reporting application or frameworks depending on our flavour.   

            