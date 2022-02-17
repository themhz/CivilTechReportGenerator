namespace ReportGenerator {
    public interface IImage {
        IImage create();
        IImage change();
        IImage delete();
        IImage copy();
        IImage paste();

    }
}