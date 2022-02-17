namespace ReportGenerator {
    public interface IListElement {
        IListElement create();
        IListElement change();
        IListElement delete();
        IListElement copy();
        IListElement paste();

        IIListElementItem listElementItem();
    }
}