namespace ReportGenerator.Handlers {
    public interface IListItem {
        int count();
        void create();
        void delete(int index);
    }
}