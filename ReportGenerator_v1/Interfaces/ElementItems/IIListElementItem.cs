namespace ReportGenerator {
    public interface IIListElementItem {
        IIListElementItem create();
        IIListElementItem add();
        IIListElementItem change();
        IIListElementItem delete();
        IIListElementItem copy();
        IIListElementItem paste();
    }
}