using System;
using System.Data;
using System.Text;
using OfficeOpenXml;
using Tool_TrainingGPT.cs;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace demo
{
    internal static class Program
    {
        enum TITLE
        {
            QUESTION = 0,
            ACTION,
            ADDRESS_NUMBER,
            ADDRESS_STREET,
            CATEGORY,
            MONTH,
            SENTENCE_OBJECT,
            SENTENCE_SUBJECT,
            SENTENCE_VERB,
            USER_ADDRESS,
            USER_CODE,
            USER_NAME,
            USER_PHONE,
            WATCH_CODE,
            WATCH_INDEX,
            YEAR
        }

        public static List<KeyValuePair<int, string>> listTitles = new List<KeyValuePair<int, string>>() {
            new KeyValuePair<int, string>((int)TITLE.ACTION, "HÀNH ĐỘNG"),
            new KeyValuePair<int, string>((int)TITLE.ADDRESS_NUMBER, "SỐ NHÀ"),
            new KeyValuePair<int, string>((int)TITLE.ADDRESS_STREET, "TÊN ĐƯỜNG"),
            new KeyValuePair<int, string>((int)TITLE.CATEGORY, "PHÂN LOẠI"),
            new KeyValuePair<int, string>((int)TITLE.MONTH, "THÁNG"),
            new KeyValuePair<int, string>((int)TITLE.QUESTION, "CÂU HỎI KHÁCH HÀNG"),
            new KeyValuePair<int, string>((int)TITLE.SENTENCE_OBJECT, "BỔ NGƯ"),
            new KeyValuePair<int, string>((int)TITLE.SENTENCE_SUBJECT, "CHỦ NGỮ"),
            new KeyValuePair<int, string>((int)TITLE.SENTENCE_VERB, "ĐỘNG TỪ"),
            new KeyValuePair<int, string>((int)TITLE.USER_ADDRESS, "ĐỊA CHỈ"),
            new KeyValuePair<int, string>((int)TITLE.USER_CODE, "MÃ KHÁCH HÀNG"),
            new KeyValuePair<int, string>((int)TITLE.USER_NAME, "TÊN KHÁCH HÀNG"),
            new KeyValuePair<int, string>((int)TITLE.USER_PHONE, "SỐ ĐIỆN THOẠI"),
            new KeyValuePair<int, string>((int)TITLE.WATCH_CODE, "MÃ ĐỒNG HỒ"),
            new KeyValuePair<int, string>((int)TITLE.WATCH_INDEX, "CHỈ SỐ NƯỚC"),
            new KeyValuePair<int, string>((int)TITLE.YEAR, "NĂM")
        };

        private static void make_file_Click()
        {
            string configFile = "name.cfg";
            string botName = "Chatbot";
            try
            {
                if (File.Exists(configFile))
                {
                    using (StreamReader reader = new StreamReader(configFile))
                    {
                        string content = reader.ReadToEnd();
                        if (!(String.IsNullOrEmpty(content) || String.IsNullOrWhiteSpace(content)))
                        {
                            botName = content;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("[ERROR] Error: " + ex.Message);
            }
            Console.WriteLine("[INFO] Chat system name: \"" + botName.Trim() + "\"");

            string excelFile = "data.xlsx";
            try
            {
                string[] excelFiles = Directory.GetFiles(".", "*.*", SearchOption.AllDirectories)
                                        .Where(file => file.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                                                       file.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                                        .ToArray();
                foreach (string file in excelFiles)
                {
                    if (!(String.IsNullOrEmpty(file) || String.IsNullOrWhiteSpace(file)))
                    {
                        excelFile = file;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("[ERROR] Error: " + ex.Message);
            }
            if (!File.Exists(excelFile))
            {
                Console.WriteLine("[ERROR] Could not find any excel file.");
                return;
            }
            Console.WriteLine("[INFO] Source excel file: \"" + excelFile + "\"");

            string filePath = "generated_training_data-" + excelFile.Trim().Replace(".\\", "").Replace(".", "_").Replace(" ", "_") + ".jsonl";
            string fileCode = "generated_code-" + excelFile.Trim().Replace(".\\", "").Replace(".", "_").Replace(" ", "_") + ".json";
            string system_role = "system";
            string user_role = "user";
            string assistant_role = "assistant";
            string gpt_content_role_system = "You are " + botName.Trim() + ". Please interpret the following user input and convert it into JSON of the form %THIS_IS_REPLACEMENT_01%. Only return JSON. User input:";
            string gpt_content_role_assistant = "%THIS_IS_REPLACEMENT_02%";

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            if (File.Exists(fileCode))
            {
                File.Delete(fileCode);
            }
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(excelFile))
                {
                    List<ActionData> actionArray = new List<ActionData>();
                    ExcelWorksheet? worksheetAction = package.Workbook.Worksheets["Type"];
                    if (worksheetAction != null)
                    {
                        int rowCount = worksheetAction.Dimension?.Rows ?? 0;
                        int colCount = worksheetAction.Dimension?.Columns ?? 0;
                        for (int i = 2; i <= colCount; i++)
                        {
                            string contentCell = worksheetAction.Cells[1, i].Text.Trim();
                            ActionData actionData = new ActionData(contentCell, i - 1);
                            for (int j = 2; j <= rowCount; j++)
                            {
                                contentCell = worksheetAction.Cells[j, i].Text.Trim();
                                if (String.IsNullOrEmpty(contentCell) || String.IsNullOrWhiteSpace(contentCell))
                                {
                                    break;
                                }
                                TypeData typeData = new TypeData(contentCell, j - 1);
                                typeData.set_action_index(actionData.get_action_index() + typeData.get_action_index());
                                actionData.list_type.Add(typeData);
                            }
                            actionArray.Add(actionData);
                        }
                    }
                    using (StreamWriter writer = new StreamWriter(fileCode, true, new UTF8Encoding(true)))
                    {
                        writer.WriteLine(Newtonsoft.Json.JsonConvert.SerializeObject(actionArray));
                    }

                    ExcelWorksheet? worksheet = package.Workbook.Worksheets["Data"];
                    if (worksheet != null)
                    {
                        Mapper mapTitle = new Mapper();
                        int rowCount = worksheet.Dimension?.Rows ?? 0;
                        int colCount = worksheet.Dimension?.Columns ?? 0;
                        Console.WriteLine("[INFO] Number of records: " + rowCount);

                        using (StreamWriter writer = new StreamWriter(filePath, true, new UTF8Encoding(true)))
                        {
                            for (int col = 1; col <= colCount; col++)
                            {
                                string? tmpStr = worksheet.Cells[1, col]?.Text?.Trim();
                                int flag = -1;
                                if (String.IsNullOrEmpty(tmpStr)) { continue; }
                                bool isFound = false;
                                foreach (var item in listTitles)
                                {
                                    if (String.Equals(item.Value, tmpStr)) { isFound = true; flag = item.Key; break; }
                                }
                                if (isFound) { mapTitle.AddTitle(flag, tmpStr, col); }
                            }

                            List<int> colArray = mapTitle.GetArrayIndex();
                            List<int> flagArray = mapTitle.GetArrayFlag();

                            for (int row = 2; row <= rowCount; row++)
                            {
                                MessageModel roleAssistant = new MessageModel();
                                MessageModel roleSystem = new MessageModel();
                                MessageModel roleUser = new MessageModel();

                                DataStructure dataStructure = new DataStructure();
                                string tmpSystemString = Regex.Unescape(JsonSerializer.Serialize(dataStructure));
                                tmpSystemString = tmpSystemString.Replace("\"", "\\\"");

                                roleAssistant.role = assistant_role;
                                roleSystem.role = system_role;
                                roleUser.role = user_role;

                                roleAssistant.content = gpt_content_role_assistant;
                                roleSystem.content = gpt_content_role_system;

                                bool isRowEmpty = false;
                                ActionData actionData = new ActionData();
                                for (int i = 0; i < colArray.Count; i++)
                                {
                                    if (isRowEmpty) break;
                                    string contentCell = worksheet.Cells[row, colArray.ElementAt(i)].Text.Trim();
                                    if (String.IsNullOrEmpty(contentCell)) { contentCell = String.Empty; }
                                    switch ((TITLE)flagArray.ElementAt(i))
                                    {
                                        case TITLE.QUESTION:
                                            if (String.IsNullOrEmpty(contentCell))
                                            {
                                                isRowEmpty = true;
                                            }
                                            else
                                            {
                                                roleUser.content = contentCell;
                                            }
                                            break;
                                        case TITLE.ACTION:
                                            if (String.IsNullOrEmpty(contentCell))
                                            {
                                                dataStructure.action = contentCell;
                                            }
                                            else
                                            {
                                                actionData = actionArray.Find(e => e.action_string == contentCell);
                                                dataStructure.action = actionData.action_index;
                                            }
                                            break;
                                        case TITLE.CATEGORY:
                                            if (String.IsNullOrEmpty(contentCell))
                                            {
                                                dataStructure.action = contentCell;
                                            }
                                            else
                                            {
                                                TypeData typeData = actionData.list_type.Find(e => e.type_string == contentCell);
                                                dataStructure.category = typeData.type_index;
                                            }
                                            break;
                                        case TITLE.USER_CODE:
                                            dataStructure.user_id = contentCell; break;
                                        case TITLE.USER_PHONE:
                                            dataStructure.phone_number = contentCell; break;
                                        case TITLE.USER_NAME:
                                            dataStructure.user_name = contentCell; break;
                                        case TITLE.USER_ADDRESS:
                                            dataStructure.address = contentCell; break;
                                        case TITLE.WATCH_CODE:
                                            dataStructure.watch_id = contentCell; break;
                                        case TITLE.WATCH_INDEX:
                                            dataStructure.watch_index_value = contentCell; break;
                                        case TITLE.MONTH:
                                            dataStructure.month = contentCell; break;
                                        case TITLE.YEAR:
                                            dataStructure.year = contentCell; break;
                                        case TITLE.ADDRESS_NUMBER:
                                            dataStructure.number_address = contentCell; break;
                                        case TITLE.ADDRESS_STREET:
                                            dataStructure.name_street = contentCell; break;
                                        case TITLE.SENTENCE_SUBJECT:
                                            dataStructure.sentence_subject = contentCell; break;
                                        case TITLE.SENTENCE_VERB:
                                            dataStructure.sentence_verb = contentCell; break;
                                        case TITLE.SENTENCE_OBJECT:
                                            dataStructure.sentence_object = contentCell; break;
                                    }
                                }

                                string tmpDataString = Regex.Unescape(JsonSerializer.Serialize(dataStructure));
                                tmpDataString = tmpDataString.Replace("\"", "\\\"");


                                ListMessages listMessages = new ListMessages();
                                listMessages.messages.Add(roleSystem);
                                listMessages.messages.Add(roleUser);
                                listMessages.messages.Add(roleAssistant);

                                string serializationString = JsonSerializer.Serialize(listMessages);
                                serializationString = Regex.Unescape(serializationString);
                                serializationString = serializationString.Replace(gpt_content_role_assistant, tmpDataString);
                                serializationString = serializationString.Replace("%THIS_IS_REPLACEMENT_01%", tmpSystemString);
                                if (!isRowEmpty)
                                {
                                    writer.WriteLine(serializationString);
                                }
                            }

                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("[ERROR] Error: " + ex.Message);
            }

            Console.WriteLine("[INFO] Create a successful training file: \"" + filePath + "\"");
        }
        private static void Main(string[] args)
        {
            make_file_Click();
            return;
        }

    }
}