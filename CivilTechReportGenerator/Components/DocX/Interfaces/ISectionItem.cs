﻿namespace CivilTechReportGenerator.Handlers {
    interface ISectionItem {
        string text { get; set; }

        int count();
        void create();
        void delete(int index);
        void replace(string generatedfile, int pos);
    }
}