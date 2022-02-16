namespace ReportGenerator {
    public interface ITableElement {
        ITableElement create();
        ITableElement change();
        ITableElement delete();
        ITableElement copy();
        ITableElement paste();
        ITableElement tableElementItem();
    }
}