namespace ReportGenerator {
    public interface ITitleElement {
        ITitleElement create();
        ITitleElement change();
        ITitleElement delete();
        ITitleElement copy();
        ITitleElement paste();
    }
}