namespace ReportGenerator {
    public interface IReport {
        IReport create();
        void delete();
        ITitleElement title();
        ITableElement table();
        IListElement list();        
        IParagraphElement paragraph();
        IImage image();
    }
}