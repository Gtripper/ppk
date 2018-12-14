using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ppk5_v2
{
    class Fabric : IFabric
    {
        private string excelPath;
        private string driverPath;
        private int numOfThreads;
        private int threadLenght;
        dynamic data;

        public Fabric(string ExcelPath, string DriverPath, int NumOfThreads, int ThreadLenght)
        {
            excelPath = ExcelPath;
            driverPath = DriverPath;
            numOfThreads = NumOfThreads;
            threadLenght = ThreadLenght;
        }

        public void SearchOKS(string columnName, int startReadInxed)
        {
            try
            {
                ExcelAppReadData read = new ExcelAppReadData(excelPath);
                read.Run("A", 2);
                MultiThread threads = new MultiThread(read._elems, numOfThreads, threadLenght, driverPath);
                threads.ThreadMaster();
                data = threads.output;
            }
            finally
            {
                ExcelAppWriteData write = new ExcelAppWriteData(data);
                write.Run();
            }
        }
    }

    interface IFabric
    {
        void SearchOKS(string columnName, int startReadIndex);
    }

    
}
