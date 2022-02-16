namespace ReportGenerator {
    public interface IParagraphElement {
        IParagraphElement create();
        IParagraphElement change();
        IParagraphElement delete();
        IParagraphElement copy();
        IParagraphElement paste();
    }
}