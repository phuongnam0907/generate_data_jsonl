using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tool_TrainingGPT.cs
{
    public class DataStructure
    {
        public string action { get; set; }
        public string action_en { get; set; }
        public string category { get; set; }
        public string category_main { get; set; }
        public string category_sub { get; set; }
        public string month { get; set; }
        public string sentence_object { get; set; }
        public string sentence_subject { get; set; }
        public string sentence_verb { get; set; }
        public string user_address { get; set; }
        public string user_address_number { get; set; }
        public string user_address_street { get; set; }
        public string user_code { get; set; }
        public string user_company { get; set; }
        public string user_name { get; set; }
        public string user_phone { get; set; }
        public string watch_code { get; set; }
        public string watch_index { get; set; }
        public string year { get; set; }

        public DataStructure()
        {
            this.action = String.Empty;
            this.action_en = String.Empty;
            this.category = String.Empty;
            this.category_main = String.Empty;
            this.category_sub = String.Empty;
            this.month = String.Empty;
            this.sentence_object = String.Empty;
            this.sentence_object = String.Empty;
            this.sentence_subject = String.Empty;
            this.sentence_verb = String.Empty;
            this.user_address = String.Empty;
            this.user_address_number = String.Empty;
            this.user_address_street = String.Empty;
            this.user_code = String.Empty;
            this.user_company = String.Empty;
            this.user_name = String.Empty;
            this.user_phone = String.Empty;
            this.watch_code = String.Empty;
            this.watch_index = String.Empty;
            this.year = String.Empty;
        }
    }
}
