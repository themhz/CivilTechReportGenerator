using System;

namespace ReportGenerator {
    public interface IDocument
    {
        void BeginScan();
        IPosition Next(); // Next element
        IPosition Next(string pattern); // Next element that contains pattern
        //
        bool SetValue(IPosition position, string value);
        bool Copy(IRange range, IPosition position);
        bool Delete(IRange range);
    }
}
