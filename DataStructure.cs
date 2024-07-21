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
        public string address { get; set; }
        public string category { get; set; }
        public string category_main { get; set; }
        public string category_sub { get; set; }
        public string device_id { get; set; }
        public string index_value { get; set; }
        public string month { get; set; }
        public string name_street { get; set; }
        public string number_address { get; set; }
        public string phone_number { get; set; }
        public string sentence_object { get; set; }
        public string sentence_subject { get; set; }
        public string sentence_verb { get; set; }
        public string user_company { get; set; }
        public string user_company_vietnamese { get; set; }
        public string user_id { get; set; }
        public string user_name { get; set; }
        public string year { get; set; }

        public DataStructure()
        {
            this.action = String.Empty;
            this.action_en = String.Empty;
            this.address = String.Empty;
            this.category = String.Empty;
            this.category_main = String.Empty;
            this.category_sub = String.Empty;
            this.device_id = String.Empty;
            this.index_value = String.Empty;
            this.month = String.Empty;
            this.name_street = String.Empty;
            this.number_address = String.Empty;
            this.phone_number = String.Empty;
            this.sentence_object = String.Empty;
            this.sentence_object = String.Empty;
            this.sentence_subject = String.Empty;
            this.sentence_verb = String.Empty;
            this.user_company = "boolean";
            this.user_company_vietnamese = "boolean";
            this.user_id = String.Empty;
            this.user_name = String.Empty;
            this.year = String.Empty;
        }
    }
}
