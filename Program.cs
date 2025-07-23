using System.Collections.Generic;

namespace ConsoleApp5
{
    internal class Program
    {
        static void Main(string[] args)
        {
            
            Console.WriteLine("Enter the size of an array");
            uint[] inparr=new uint[5] {1,5,8,4,6};
            int num = 5;
            List<List<int>> retlist = calcfun( inparr, num);
            foreach (var i in retlist) {
                Console.WriteLine($"The first index is {i[0]} and second index is {i[1]}");
            }
        }
        public static List<List<int>> calcfun(uint[] inparr,int num)
        {
            List<List<int>> list = new List<List<int>>();
            for (int i = 0; i < inparr.Length; i++)
            {
                for (int j = 0; j < inparr.Length; j++)
                {
                    if (i == j)
                    {
                        continue;
                    }
                    else
                    {
                        if (inparr[i] + inparr[j] == num)
                        {
                            List<int> locallist = new List<int>();
                            locallist.Add(i);
                            locallist.Add(j);
                            list.Add(locallist);
                        }

                    }
                }
            }
            return list;
        }
    }
}
