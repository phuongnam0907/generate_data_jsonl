using System;
using System.Collections;
using System.Data;
using System.Reflection.Metadata;
using System.Text;
using System.Collections.Generic;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using System.ComponentModel.Design;
using Tool_TrainingGPT.cs;
using static System.Collections.Specialized.BitVector32;
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
            CATEGORY,
            USER_CODE,
            USER_PHONE,
            USER_NAME,
            USER_ADDRESS,
            WATCH_CODE,
            WATCH_INDEX,
            MONTH,
            YEAR,
            ADDRESS_NUMBER,
            ADDRESS_STREET,
            USER_COMPANY,
            ACTION_ENGLISH,
            CATEGORY_MAIN,
            CATEGORY_SUB,
            SENTENCE_SUBJECT,
            SENTENCE_VERB,
            SENTENCE_OBJECT
        }

        public static List<KeyValuePair<int, string>> listTitles = new List<KeyValuePair<int, string>>() {
            new KeyValuePair<int, string>((int)TITLE.QUESTION, "CÂU HỎI KHÁCH HÀNG"),
            new KeyValuePair<int, string>((int)TITLE.ACTION, "HÀNH ĐỘNG"),
            new KeyValuePair<int, string>((int)TITLE.CATEGORY, "PHÂN LOẠI"),
            new KeyValuePair<int, string>((int)TITLE.USER_CODE, "MÃ KHÁCH HÀNG"),
            new KeyValuePair<int, string>((int)TITLE.USER_PHONE, "SỐ ĐIỆN THOẠI"),
            new KeyValuePair<int, string>((int)TITLE.USER_NAME, "TÊN KHÁCH HÀNG"),
            new KeyValuePair<int, string>((int)TITLE.USER_ADDRESS, "ĐỊA CHỈ"),
            new KeyValuePair<int, string>((int)TITLE.WATCH_CODE, "MÃ ĐỒNG HỒ"),
            new KeyValuePair<int, string>((int)TITLE.WATCH_INDEX, "CHỈ SỐ NƯỚC"),
            new KeyValuePair<int, string>((int)TITLE.MONTH, "THÁNG"),
            new KeyValuePair<int, string>((int)TITLE.YEAR, "NĂM"),
            new KeyValuePair<int, string>((int)TITLE.ADDRESS_NUMBER, "SỐ NHÀ"),
            new KeyValuePair<int, string>((int)TITLE.ADDRESS_STREET, "TÊN ĐƯỜNG"),
            new KeyValuePair<int, string>((int)TITLE.USER_COMPANY, "CƠ QUAN_DOANG NGHIỆP"),
            new KeyValuePair<int, string>((int)TITLE.ACTION_ENGLISH, "English"),
            new KeyValuePair<int, string>((int)TITLE.CATEGORY_MAIN, "Phân Loại 1"),
            new KeyValuePair<int, string>((int)TITLE.CATEGORY_SUB, "Phân Loại 2"),
            new KeyValuePair<int, string>((int)TITLE.SENTENCE_SUBJECT, "Chủ ngữ"),
            new KeyValuePair<int, string>((int)TITLE.SENTENCE_VERB, "Động từ"),
            new KeyValuePair<int, string>((int)TITLE.SENTENCE_OBJECT, "Bổ ngữ")
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

            string filePath = "training_data-" + excelFile.Trim().Replace(".\\", "").Replace(".", "_").Replace(" ", "_") + ".jsonl";
            string system_role = "system";
            string user_role = "user";
            string assistant_role = "assistant";
            string gpt_content_role_system = "You are " + botName.Trim() + ". Please interpret the following user input and convert it into JSON of the form %THIS_IS_REPLACEMENT_01%. Only return JSON. User input:";
            string gpt_content_role_assistant = "%THIS_IS_REPLACEMENT_02%";

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(excelFile))
                {
                    ExcelWorksheet? worksheet = package.Workbook.Worksheets.FirstOrDefault();
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
                                string tmpStr = worksheet.Cells[1, col]?.Text?.Trim();
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

                                for (int i = 0; i < colArray.Count; i++)
                                {
                                    string contentCell = worksheet.Cells[row, colArray.ElementAt(i)].Text.Trim();
                                    if (String.IsNullOrEmpty(contentCell)) { contentCell = String.Empty; }
                                    switch ((TITLE)flagArray.ElementAt(i))
                                    {
                                        case TITLE.QUESTION:
                                            roleUser.content = contentCell; break;
                                        case TITLE.ACTION:
                                            dataStructure.action = contentCell; break;
                                        case TITLE.CATEGORY:
                                            dataStructure.category = contentCell; break;
                                        case TITLE.USER_CODE:
                                            dataStructure.user_code = contentCell; break;
                                        case TITLE.USER_PHONE:
                                            dataStructure.user_phone = contentCell; break;
                                        case TITLE.USER_NAME:
                                            dataStructure.user_name = contentCell; break;
                                        case TITLE.USER_ADDRESS:
                                            dataStructure.user_address = contentCell; break;
                                        case TITLE.WATCH_CODE:
                                            dataStructure.watch_code = contentCell; break;
                                        case TITLE.WATCH_INDEX:
                                            dataStructure.watch_index = contentCell; break;
                                        case TITLE.MONTH:
                                            dataStructure.month = contentCell; break;
                                        case TITLE.YEAR:
                                            dataStructure.year = contentCell; break;
                                        case TITLE.ADDRESS_NUMBER:
                                            dataStructure.user_address_number = contentCell; break;
                                        case TITLE.ADDRESS_STREET:
                                            dataStructure.user_address_street = contentCell; break;
                                        case TITLE.USER_COMPANY:
                                            if (String.IsNullOrEmpty(contentCell)) contentCell = "false";
                                            dataStructure.user_company = contentCell; break;
                                        case TITLE.ACTION_ENGLISH:
                                            dataStructure.action_en = contentCell; break;
                                        case TITLE.CATEGORY_MAIN:
                                            dataStructure.category_main = contentCell; break;
                                        case TITLE.CATEGORY_SUB:
                                            dataStructure.category_sub = contentCell; break;
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
                                writer.WriteLine(serializationString);
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