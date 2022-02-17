using ReportGenerator.Interfaces.Elements;

namespace ReportGenerator {
    public interface ITableElement {
        ITableElement setTableIndex(int index);
        ITableElement create();        
        ITableElement delete();
        ITableElement copy(int index);
        ITableElement paste();
        ITableElement tableElementItem();
        int count();
        
    }
}