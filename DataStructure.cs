using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Tool_TrainingGPT.cs
{
    public class DataStructure
    {
        public string action { get; set; }
        public string address { get; set; }
        public string category { get; set; }
        public string month { get; set; }
        public string name_street { get; set; }
        public string number_address { get; set; }
        public string phone_number { get; set; }
        public string sentence_object { get; set; }
        public string sentence_subject { get; set; }
        public string sentence_verb { get; set; }
        public string user_id { get; set; }
        public string user_name { get; set; }
        public string watch_id { get; set; }
        public string watch_index_value { get; set; }
        public string year { get; set; }

        public DataStructure()
        {
            this.action = String.Empty;
            this.address = String.Empty;
            this.category = String.Empty;
            this.month = String.Empty;
            this.name_street = String.Empty;
            this.number_address = String.Empty;
            this.phone_number = String.Empty;
            this.sentence_object = String.Empty;
            this.sentence_subject = String.Empty;
            this.sentence_verb = String.Empty;
            this.user_id = String.Empty;
            this.user_name = String.Empty;
            this.watch_id = String.Empty;
            this.watch_index_value = String.Empty;
            this.year = String.Empty;
        }
    }

    public class ActionData
    {
        public string action_string { get; set; }

        public string action_index;

        public string get_action_index()
        {
            return action_index;
        }

        public void set_action_index(int value)
        {
            action_index = value.ToString("000");
        }

        public List<TypeData> list_type { get; set; }

        public ActionData()
        {
            this.action_string = String.Empty;
            this.action_index = "000";
            this.list_type = new List<TypeData>();
        }

        public ActionData(string name, int index)
        {
            this.action_string = name;
            set_action_index(index);
            this.list_type = new List<TypeData>();
        }
    }

    public class TypeData
    {
        public string type_string { get; set; }

        public string type_index;

        public string get_action_index()
        {
            return type_index;
        }

        public void set_action_index(int value)
        {
            type_index = value.ToString("000");
        }

        public void set_action_index(string value)
        {
            type_index = value;
        }
        public TypeData()
        {
            this.type_string = String.Empty;
            this.type_index = "000";
        }

        public TypeData(string name, int index)
        {
            this.type_string = name;
            set_action_index(index);
        }
    }
}
