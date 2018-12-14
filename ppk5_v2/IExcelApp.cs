namespace ppk5_v2
{
    interface IExcelAppWrite
    {
        void Run();
    }

    interface IExcelAppRead
    {
        dynamic _elems();
        void Run();
        void Run(string columnName);
        void Run(string columnName, int startIndex);
        void Run(string columnName, int startIndex, int endIndex);
    }
}