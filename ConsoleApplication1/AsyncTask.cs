using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    public static class AsyncTask
    {
        public static async Task<int> CalculateAnswer()
        {
            var firstTask = OperationOneAsync();
            var secondTask = OperationTwoAsync();

            return await firstTask + await secondTask;
        }

        private static async Task<int> OperationTwoAsync()
        {
            await Task.Delay(5);
            return 2;
        }

        private static async Task<int> OperationOneAsync()
        {
            await Task.Delay(1);
            return 1;
        }
    }
}