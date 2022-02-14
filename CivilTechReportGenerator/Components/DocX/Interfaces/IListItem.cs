namespace CivilTechReportGenerator.Handlers {
    interface IListItem {
        int count();
        void create();
        void delete(int index);
    }
}