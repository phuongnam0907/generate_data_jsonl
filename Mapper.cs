using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tool_TrainingGPT.cs
{
    internal class Index
    {
        public string Title { get; set; }
        public int Indexing { get; set; }
        public int Flag { get; set; }

        public Index()
        {
            Title = string.Empty;
            Indexing = 0;
            Flag = 0;
        }
    }
    public class Mapper
    {
        private List<Index> _maps = new List<Index>();
        public void AddTitle(int flag, string title, int index) { _maps.Add(new Index { Flag = flag, Title = title, Indexing = index }); }

        public List<int> GetArrayIndex()
        {
            List<int> indexes = new List<int>();
            foreach (var map in _maps)
            {
                indexes.Add(map.Indexing);
            }
            return indexes;
        }
        public List<int> GetArrayFlag()
        {
            List<int> indexes = new List<int>();
            foreach (var map in _maps)
            {
                indexes.Add(map.Flag);
            }
            return indexes;
        }

        public void PrintAll()
        {
            foreach (var map in _maps)
            {
                Console.WriteLine(map.Indexing + " -- " + map.Title + " [" + map.Flag + "]");
            }
        }
    }
}
