using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ppk5_v2
{
    class MultiThread
    {
        private int numOfThreads;
        private int threadLenght;
        private string driverPath;
        internal IEnumerable<List<Elem>> output;

        public MultiThread(List<Elem> input, int NumOfThread, int ThreadLenght, string DriverPath)
        {
            numOfThreads = NumOfThread;
            threadLenght = ThreadLenght;
            driverPath = DriverPath;
            output = splitList(input);            
        }



        private IEnumerable<List<T>> splitList<T>(List<T> input)
        {
            for (int i = 0; i < input.Count; i += threadLenght)
            {
                yield return input.GetRange(i, Math.Min(threadLenght, input.Count - i));
            }
        }

        public void ThreadMaster()
        {
            for (int i = 0; i < output.Count(); i += numOfThreads)
            {
                // Пока число необработанный элементов больше количеств потоков
                if (output.Count() - i >= numOfThreads)
                {
                    Task[] tasks1 = new Task[numOfThreads];
                    for (var j = 0; j < tasks1.Length; j++)
                    {
                        var index = i + j;
                        tasks1[j] = Task.Factory.StartNew(() => { Parser parser = new Parser
                                (driverPath, output.ElementAt(index));
                                parser.parser(); });
                    }
                    try
                    {
                        Task.WaitAll(tasks1); // ожидаем завершения задач 
                    }
                    catch (AggregateException e)
                    {
                        foreach (var task in tasks1)
                        {
                            if (task.Status == TaskStatus.Faulted)
                            {
                                Console.WriteLine(task.Exception.GetType().BaseType.Name); 
                                Console.WriteLine(task.Exception.GetType().Name);
                                Console.WriteLine();
                            }
                        }
                    }
                }
                // Создаем потоки на оставшееся число элементов output
                else
                {
                    int N = output.Count() - i;

                    Task[] tasks2 = new Task[N];
                    for (var j = 0; j < tasks2.Length; j++)
                    {
                        var index = i + j;
                        tasks2[j] = Task.Factory.StartNew(() => { Parser parser = new Parser
                                (driverPath, output.ElementAt(index));
                                parser.parser(); });
                    }
                    Task.WaitAll(tasks2);
                }
            }
        }
    }
}
