namespace ReportGenerator {
    public interface ITableElementItem {
        ITableElementItem create();
        ITableElementItem add();
        ITableElementItem change();
        ITableElementItem delete();
        ITableElementItem copy();
        ITableElementItem paste();
    }
}