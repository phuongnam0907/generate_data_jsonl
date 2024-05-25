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

namespace demo{
    internal static class Program{
        private static void make_file_Click(){
            string excelFile = "data.xlsx";
            string filePath = "training_data.jsonl";
            
            if (File.Exists(filePath)){
                File.Delete(filePath);
            }
            bool checked_test = false;
            int countKB1 = 0; int countKB11 = 0; int countKB12 = 0;
            int countKB2 = 0; int countKB24 = 0; int countKB25 = 0; int countKB26 = 0;
            int countKB3 = 0; int countKB37 = 0; int countKB38 = 0; int countKB39 = 0;
            int countKB4 = 0; int countKB41 = 0; int countKB42 = 0;
            int countKB5 = 0; int countKB51 = 0; int countKB52 = 0;
            int countKB6 = 0; int countKB61 = 0; int countKB62 = 0;
            int countKB7 = 0; int countKB71 = 0; int countKB72 = 0;
            int countSingle = 0; int countFlag0 = 0;
            int countFlag1 = 0; int countFlag2 = 0; int countFlag3 = 0; int countFlag4 = 0; int countFlag5 = 0; int countFlag6 = 0; int countFlag7 = 0; 
            string ask_phone = "Anh/chị vui lòng cung cấp số điện thoại liên hệ";
            string ask_address = "Anh/chị vui lòng cung cấp địa chỉ của đồng hồ nước";
            string ask_device_id = "Anh/chị vui lòng cung cấp mã đồng hồ nước";
            string ask_month = "Anh/chị vui lòng cung cấp tháng mà anh/chị muốn kiểm tra";
            string ask_index_water = "Anh/Chị vui lòng nhập dãy số màu đen trên mặt đồng hồ nước";
            string ask_user_id = "Anh/chị vui lòng cung cấp mã khách hàng";
            string ask_user_name = "Anh/chị vui lòng cung cấp tên khách hàng của đồng hồ nước";
            string ask_user_name_phone_address = "Anh/chị vui lòng cung cấp địa chỉ lắp đặt đồng hồ, tên chủ hộ, số điện thoại";
            try{
                string gpt_content_role = "You are Bwaco Chatbot. Please interpret the following user input and convert it into JSON of the form " +
                   "{\\\"hành_động\\\":\\\"\\\",\\\"phân_loại\\\":\\\"\\\",\\\"mã_khách_hàng\\\":\\\"\\\",\\\"số_điện_thoại\\\":\\\"\\\",\\\"tên_khách_hàng\\\":\\\"\\\"," +
                   "\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\",\\\"chỉ_đồng_hồ\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"boolean\\\"}. Only return JSON. User input:";

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(excelFile)){
                    ExcelWorksheet? worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet != null){
                        int rowCount = worksheet.Dimension?.Rows ?? 0;
                        using (StreamWriter writer = new StreamWriter(filePath, true)){
                            //===============================================SINGLE SENTENCES============================================
                            
                            for (int row = 2; row <= rowCount; row++){
                                string? question = string.IsNullOrEmpty(worksheet.Cells[row, 2]?.Text?.Trim()) ? "" : worksheet.Cells[row, 2]?.Text?.Trim();
                                string? action = string.IsNullOrEmpty(worksheet.Cells[row, 3]?.Text?.Trim()) ? "" : worksheet.Cells[row, 3]?.Text?.Trim();
                                string? category = string.IsNullOrEmpty(worksheet.Cells[row, 4]?.Text?.Trim()) ? "" : worksheet.Cells[row, 4]?.Text?.Trim();
                                string? user_id = string.IsNullOrEmpty(worksheet.Cells[row, 5]?.Text?.Trim()) ? "" : worksheet.Cells[row, 5]?.Text?.Trim();
                                string? phone_number = string.IsNullOrEmpty(worksheet.Cells[row, 6]?.Text?.Trim()) ? "" : worksheet.Cells[row, 6]?.Text?.Trim();
                                string? user_name = string.IsNullOrEmpty(worksheet.Cells[row, 7]?.Text?.Trim()) ? "" : worksheet.Cells[row, 7]?.Text?.Trim();
                                string? address = string.IsNullOrEmpty(worksheet.Cells[row, 8]?.Text?.Trim()) ? "" : worksheet.Cells[row, 8]?.Text?.Trim();
                                string? device_id = string.IsNullOrEmpty(worksheet.Cells[row, 9]?.Text?.Trim()) ? "" : worksheet.Cells[row, 9]?.Text?.Trim();
                                string? index_water = string.IsNullOrEmpty(worksheet.Cells[row, 10]?.Text?.Trim()) ? "" : worksheet.Cells[row, 10]?.Text?.Trim();
                                string? month = string.IsNullOrEmpty(worksheet.Cells[row, 11]?.Text?.Trim()) ? "" : worksheet.Cells[row, 11]?.Text?.Trim();
                                string? year = string.IsNullOrEmpty(worksheet.Cells[row, 12]?.Text?.Trim()) ? "" : worksheet.Cells[row, 12]?.Text?.Trim();
                                string? red_flag = string.IsNullOrEmpty(worksheet.Cells[row, 13]?.Text?.Trim()) ? "" : worksheet.Cells[row, 13]?.Text?.Trim();
                                string? num_address = string.IsNullOrEmpty(worksheet.Cells[row, 14]?.Text?.Trim()) ? "" : worksheet.Cells[row, 14]?.Text?.Trim();
                                string? name_street = string.IsNullOrEmpty(worksheet.Cells[row, 15]?.Text?.Trim()) ? "" : worksheet.Cells[row, 15]?.Text?.Trim();
                                string? ca_nhan = string.IsNullOrEmpty(worksheet.Cells[row, 16]?.Text?.Trim()) ? "" : worksheet.Cells[row, 16]?.Text?.Trim();
                                if(ca_nhan == ""){
                                    ca_nhan = "false";
                                }
                                string user_content_role = "{\"role\": \"user\", \"content\":" + $"\"{question}\"" + "},";
                                string assistant_content_role = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"{action}\\\",\\\"phân_loại\\\":\\\"{category}\\\"," +
                                $"\\\"mã_khách_hàng\\\":\\\"{user_id}\\\", \\\"số_điện_thoại\\\":\\\"{phone_number}\\\", \\\"tên_khách_hàng\\\":\\\"{user_name}\\\",\\\"địa_chỉ\\\":\\\"{address}\\\",\\\"mã_đồng_hồ\\\":\\\"{device_id}\\\"," +
                                $"\\\"chỉ_số_nước\\\":\\\"{index_water}\\\",\\\"tháng\\\":\\\"{month}\\\",\\\"năm\\\":\\\"{year}\\\",\\\"số_nhà\\\":\\\"{num_address}\\\",\\\"tên_đường\\\":\\\"{name_street}\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"{ca_nhan}\\\"}}\"}}]}}";

                                string generate_text = "{\"messages\": [{ \"role\": \"system\", \"content\": " + $"\"{gpt_content_role}\"" + "}," + user_content_role + assistant_content_role;
                                //Console.WriteLine(generate_text);
                                

                                // comment to check
                                if((red_flag != "0") && (red_flag != "1") && (red_flag != "2")  && (red_flag != "3")  && (red_flag != "4") && (red_flag != "5")  && (red_flag != "6")  && 
                                   (red_flag != "7") && (red_flag != "11")&& (red_flag != "12") && (red_flag != "14") && (red_flag != "15")&& (red_flag != "16") && (red_flag != "17") &&
                                   (red_flag != "18")&& (red_flag != "19")&& (red_flag != "20") && (red_flag != "21") && (red_flag != "22") && (red_flag != "23") && (red_flag != "26")){
                                    // Write the text to a new line
                                    
                                    writer.WriteLine(generate_text);
                                    countSingle++;
                                }
                                //*///comment to check
                                

                                /* //comment to check
                                
                                // if (checked_test == true){
                                //     break;
                                // } */
                                else{
                                    if(red_flag == "0"){
                                        string gpt_content_role0 = "You are Bwaco Chatbot. Please interpret the following user input and convert it into JSON of the form " +
                                        "{\\\"hành_động\\\":\\\"\\\",\\\"phân_loại\\\":\\\"\\\",\\\"mã_khách_hàng\\\":\\\"\\\",\\\"số_điện_thoại\\\":\\\"\\\",\\\"tên_khách_hàng\\\":\\\"\\\"," +
                                        "\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\",\\\"chỉ_đồng_hồ\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"boolean\\\"}. Only return JSON. User input:";
                                        
                                        string assistant_content_role01 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"\\\",\\\"phân_loại\\\":\\\"\\\"," +
                                        $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                        $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Anh/chị vui lòng cung cấp địa chỉ của đồng hồ nước\\\"}}\"}}";
                                        
                                        string user_content_role0 = "{\"role\": \"user\", \"content\":" + $"\"{question}\"" + "},";

                                        string assistant_content_role02 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"{action}\\\",\\\"phân_loại\\\":\\\"{category}\\\"," +
                                        $"\\\"mã_khách_hàng\\\":\\\"{user_id}\\\", \\\"số_điện_thoại\\\":\\\"{phone_number}\\\", \\\"tên_khách_hàng\\\":\\\"{user_name}\\\",\\\"địa_chỉ\\\":\\\"{address}\\\",\\\"mã_đồng_hồ\\\":\\\"{device_id}\\\"," +
                                        $"\\\"chỉ_số_nước\\\":\\\"{index_water}\\\",\\\"tháng\\\":\\\"{month}\\\",\\\"năm\\\":\\\"{year}\\\",\\\"số_nhà\\\":\\\"{num_address}\\\",\\\"tên_đường\\\":\\\"{name_street}\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\"}}\"}}]}}";

                                        string generate_text0 = "{\"messages\": [{ \"role\": \"system\", \"content\": " + $"\"{gpt_content_role0}\"" + "}," + assistant_content_role01  + "," + user_content_role0 + assistant_content_role02;
                                        writer.WriteLine(generate_text0);
                                        countFlag0++;
                                    }
                                    else if(red_flag == "1"){
                                        string gpt_content_role1 = "You are Bwaco Chatbot. Please interpret the following user input and convert it into JSON of the form " +
                                        "{\\\"hành_động\\\":\\\"\\\",\\\"phân_loại\\\":\\\"\\\",\\\"mã_khách_hàng\\\":\\\"\\\",\\\"số_điện_thoại\\\":\\\"\\\",\\\"tên_khách_hàng\\\":\\\"\\\"," +
                                        "\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\",\\\"chỉ_đồng_hồ\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"boolean\\\"}. Only return JSON. User input:";
                                        
                                        string assistant_content_role11 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"\\\",\\\"phân_loại\\\":\\\"\\\"," +
                                        $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                        $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Anh/chị vui lòng cung cấp số điện thoại liên hệ\\\"}}\"}}";
                                        
                                        string user_content_role1 = "{\"role\": \"user\", \"content\":" + $"\"{question}\"" + "},";

                                        string assistant_content_role12 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"{action}\\\",\\\"phân_loại\\\":\\\"{category}\\\"," +
                                        $"\\\"mã_khách_hàng\\\":\\\"{user_id}\\\", \\\"số_điện_thoại\\\":\\\"{phone_number}\\\", \\\"tên_khách_hàng\\\":\\\"{user_name}\\\",\\\"địa_chỉ\\\":\\\"{address}\\\",\\\"mã_đồng_hồ\\\":\\\"{device_id}\\\"," +
                                        $"\\\"chỉ_số_nước\\\":\\\"{index_water}\\\",\\\"tháng\\\":\\\"{month}\\\",\\\"năm\\\":\\\"{year}\\\",\\\"số_nhà\\\":\\\"{num_address}\\\",\\\"tên_đường\\\":\\\"{name_street}\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\"}}\"}}]}}";

                                        string generate_text1 = "{\"messages\": [{ \"role\": \"system\", \"content\": " + $"\"{gpt_content_role1}\"" + "}," + assistant_content_role11  + "," + user_content_role1 + assistant_content_role12;
                                        writer.WriteLine(generate_text1);
                                        countFlag1++;
                                    }
                                    else if(red_flag == "2"){
                                        string gpt_content_role2 = "You are Bwaco Chatbot. Please interpret the following user input and convert it into JSON of the form " +
                                        "{\\\"hành_động\\\":\\\"\\\",\\\"phân_loại\\\":\\\"\\\",\\\"mã_khách_hàng\\\":\\\"\\\",\\\"số_điện_thoại\\\":\\\"\\\",\\\"tên_khách_hàng\\\":\\\"\\\"," +
                                        "\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\",\\\"chỉ_đồng_hồ\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"boolean\\\"}. Only return JSON. User input:";
                                        
                                        string assistant_content_role21 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"\\\",\\\"phân_loại\\\":\\\"\\\"," +
                                        $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                        $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Anh/chị vui lòng cung cấp mã đồng hồ nước\\\"}}\"}}";
                                        
                                        string user_content_role2 = "{\"role\": \"user\", \"content\":" + $"\"{question}\"" + "},";

                                        string assistant_content_role22 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"{action}\\\",\\\"phân_loại\\\":\\\"{category}\\\"," +
                                        $"\\\"mã_khách_hàng\\\":\\\"{user_id}\\\", \\\"số_điện_thoại\\\":\\\"{phone_number}\\\", \\\"tên_khách_hàng\\\":\\\"{user_name}\\\",\\\"địa_chỉ\\\":\\\"{address}\\\",\\\"mã_đồng_hồ\\\":\\\"{device_id}\\\"," +
                                        $"\\\"chỉ_số_nước\\\":\\\"{index_water}\\\",\\\"tháng\\\":\\\"{month}\\\",\\\"năm\\\":\\\"{year}\\\",\\\"số_nhà\\\":\\\"{num_address}\\\",\\\"tên_đường\\\":\\\"{name_street}\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\"}}\"}}]}}";

                                        string generate_text2 = "{\"messages\": [{ \"role\": \"system\", \"content\": " + $"\"{gpt_content_role2}\"" + "}," + assistant_content_role21  + "," + user_content_role2 + assistant_content_role22;
                                        writer.WriteLine(generate_text2);
                                        countFlag2++;
                                    }
                                    else if(red_flag == "3"){
                                        string gpt_content_role3 = "You are Bwaco Chatbot. Please interpret the following user input and convert it into JSON of the form " +
                                        "{\\\"hành_động\\\":\\\"\\\",\\\"phân_loại\\\":\\\"\\\",\\\"mã_khách_hàng\\\":\\\"\\\",\\\"số_điện_thoại\\\":\\\"\\\",\\\"tên_khách_hàng\\\":\\\"\\\"," +
                                        "\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\",\\\"chỉ_đồng_hồ\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"boolean\\\"}. Only return JSON. User input:";
                                        
                                        string assistant_content_role31 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"\\\",\\\"phân_loại\\\":\\\"\\\"," +
                                        $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                        $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Anh/chị vui lòng cung cấp tháng mà anh/chị muốn kiểm tra\\\"}}\"}}";
                                        
                                        string user_content_role3 = "{\"role\": \"user\", \"content\":" + $"\"{question}\"" + "},";

                                        string assistant_content_role32 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"{action}\\\",\\\"phân_loại\\\":\\\"{category}\\\"," +
                                        $"\\\"mã_khách_hàng\\\":\\\"{user_id}\\\", \\\"số_điện_thoại\\\":\\\"{phone_number}\\\", \\\"tên_khách_hàng\\\":\\\"{user_name}\\\",\\\"địa_chỉ\\\":\\\"{address}\\\",\\\"mã_đồng_hồ\\\":\\\"{device_id}\\\"," +
                                        $"\\\"chỉ_số_nước\\\":\\\"{index_water}\\\",\\\"tháng\\\":\\\"{month}\\\",\\\"năm\\\":\\\"{year}\\\",\\\"số_nhà\\\":\\\"{num_address}\\\",\\\"tên_đường\\\":\\\"{name_street}\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\"}}\"}}]}}";

                                        string generate_text3 = "{\"messages\": [{ \"role\": \"system\", \"content\": " + $"\"{gpt_content_role3}\"" + "}," + assistant_content_role31  + "," + user_content_role3 + assistant_content_role32;
                                        writer.WriteLine(generate_text3);
                                        countFlag3++;
                                    }
                                    else if(red_flag == "4"){
                                        string gpt_content_role4 = "You are Bwaco Chatbot. Please interpret the following user input and convert it into JSON of the form " +
                                        "{\\\"hành_động\\\":\\\"\\\",\\\"phân_loại\\\":\\\"\\\",\\\"mã_khách_hàng\\\":\\\"\\\",\\\"số_điện_thoại\\\":\\\"\\\",\\\"tên_khách_hàng\\\":\\\"\\\"," +
                                        "\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\",\\\"chỉ_đồng_hồ\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"boolean\\\"}. Only return JSON. User input:";
                                        
                                        string assistant_content_role41 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"\\\",\\\"phân_loại\\\":\\\"\\\"," +
                                        $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                        $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Anh/Chị vui lòng nhập dãy số màu đen trên mặt đồng hồ nước\\\"}}\"}}";
                                        
                                        string user_content_role4 = "{\"role\": \"user\", \"content\":" + $"\"{question}\"" + "},";

                                        string assistant_content_role42 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"{action}\\\",\\\"phân_loại\\\":\\\"{category}\\\"," +
                                        $"\\\"mã_khách_hàng\\\":\\\"{user_id}\\\", \\\"số_điện_thoại\\\":\\\"{phone_number}\\\", \\\"tên_khách_hàng\\\":\\\"{user_name}\\\",\\\"địa_chỉ\\\":\\\"{address}\\\",\\\"mã_đồng_hồ\\\":\\\"{device_id}\\\"," +
                                        $"\\\"chỉ_số_nước\\\":\\\"{index_water}\\\",\\\"tháng\\\":\\\"{month}\\\",\\\"năm\\\":\\\"{year}\\\",\\\"số_nhà\\\":\\\"{num_address}\\\",\\\"tên_đường\\\":\\\"{name_street}\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\"}}\"}}]}}";

                                        string generate_text4 = "{\"messages\": [{ \"role\": \"system\", \"content\": " + $"\"{gpt_content_role4}\"" + "}," + assistant_content_role41  + "," + user_content_role4 + assistant_content_role42;
                                        writer.WriteLine(generate_text4);
                                        countFlag4++;
                                    }
                                    else if(red_flag == "5"){
                                        string gpt_content_role5 = "You are Bwaco Chatbot. Please interpret the following user input and convert it into JSON of the form " +
                                        "{\\\"hành_động\\\":\\\"\\\",\\\"phân_loại\\\":\\\"\\\",\\\"mã_khách_hàng\\\":\\\"\\\",\\\"số_điện_thoại\\\":\\\"\\\",\\\"tên_khách_hàng\\\":\\\"\\\"," +
                                        "\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\",\\\"chỉ_đồng_hồ\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"boolean\\\"}. Only return JSON. User input:";
                                        
                                        string assistant_content_role51 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"\\\",\\\"phân_loại\\\":\\\"\\\"," +
                                        $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                        $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Anh/chị vui lòng cung cấp mã khách hàng\\\"}}\"}}";
                                        
                                        string user_content_role5 = "{\"role\": \"user\", \"content\":" + $"\"{question}\"" + "},";

                                        string assistant_content_role52 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"{action}\\\",\\\"phân_loại\\\":\\\"{category}\\\"," +
                                        $"\\\"mã_khách_hàng\\\":\\\"{user_id}\\\", \\\"số_điện_thoại\\\":\\\"{phone_number}\\\", \\\"tên_khách_hàng\\\":\\\"{user_name}\\\",\\\"địa_chỉ\\\":\\\"{address}\\\",\\\"mã_đồng_hồ\\\":\\\"{device_id}\\\"," +
                                        $"\\\"chỉ_số_nước\\\":\\\"{index_water}\\\",\\\"tháng\\\":\\\"{month}\\\",\\\"năm\\\":\\\"{year}\\\",\\\"số_nhà\\\":\\\"{num_address}\\\",\\\"tên_đường\\\":\\\"{name_street}\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\"}}\"}}]}}";

                                        string generate_text5 = "{\"messages\": [{ \"role\": \"system\", \"content\": " + $"\"{gpt_content_role5}\"" + "}," + assistant_content_role51  + "," + user_content_role5 + assistant_content_role52;
                                        writer.WriteLine(generate_text5);
                                        countFlag5++;
                                    }
                                    else if(red_flag == "6"){
                                        string gpt_content_role6 = "You are Bwaco Chatbot. Please interpret the following user input and convert it into JSON of the form " +
                                        "{\\\"hành_động\\\":\\\"\\\",\\\"phân_loại\\\":\\\"\\\",\\\"mã_khách_hàng\\\":\\\"\\\",\\\"số_điện_thoại\\\":\\\"\\\",\\\"tên_khách_hàng\\\":\\\"\\\"," +
                                        "\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\",\\\"chỉ_đồng_hồ\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"boolean\\\"}. Only return JSON. User input:";
                                        
                                        string assistant_content_role61 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"\\\",\\\"phân_loại\\\":\\\"\\\"," +
                                        $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                        $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Anh/chị vui lòng cung cấp tên khách hàng\\\"}}\"}}";
                                        
                                        string user_content_role6 = "{\"role\": \"user\", \"content\":" + $"\"{question}\"" + "},";

                                        string assistant_content_role62 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"{action}\\\",\\\"phân_loại\\\":\\\"{category}\\\"," +
                                        $"\\\"mã_khách_hàng\\\":\\\"{user_id}\\\", \\\"số_điện_thoại\\\":\\\"{phone_number}\\\", \\\"tên_khách_hàng\\\":\\\"{user_name}\\\",\\\"địa_chỉ\\\":\\\"{address}\\\",\\\"mã_đồng_hồ\\\":\\\"{device_id}\\\"," +
                                        $"\\\"chỉ_số_nước\\\":\\\"{index_water}\\\",\\\"tháng\\\":\\\"{month}\\\",\\\"năm\\\":\\\"{year}\\\",\\\"số_nhà\\\":\\\"{num_address}\\\",\\\"tên_đường\\\":\\\"{name_street}\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\"}}\"}}]}}";

                                        string generate_text6 = "{\"messages\": [{ \"role\": \"system\", \"content\": " + $"\"{gpt_content_role6}\"" + "}," + assistant_content_role61  + "," + user_content_role6 + assistant_content_role62;
                                        writer.WriteLine(generate_text6);
                                        countFlag6++;
                                    }
                                    else if(red_flag == "7"){
                                        string gpt_content_role7 = "You are Bwaco Chatbot. Please interpret the following user input and convert it into JSON of the form " +
                                        "{\\\"hành_động\\\":\\\"\\\",\\\"phân_loại\\\":\\\"\\\",\\\"mã_khách_hàng\\\":\\\"\\\",\\\"số_điện_thoại\\\":\\\"\\\",\\\"tên_khách_hàng\\\":\\\"\\\"," +
                                        "\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\",\\\"chỉ_đồng_hồ\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"boolean\\\"}. Only return JSON. User input:";
                                        
                                        string assistant_content_role71 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"\\\",\\\"phân_loại\\\":\\\"\\\"," +
                                        $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                        $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Anh/chị vui lòng cung cấp địa chỉ lắp đặt đồng hồ, tên chủ hộ, số điện thoại\\\"}}\"}}";
                                        
                                        string user_content_role7 = "{\"role\": \"user\", \"content\":" + $"\"{question}\"" + "},";

                                        string assistant_content_role72 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"{action}\\\",\\\"phân_loại\\\":\\\"{category}\\\"," +
                                        $"\\\"mã_khách_hàng\\\":\\\"{user_id}\\\", \\\"số_điện_thoại\\\":\\\"{phone_number}\\\", \\\"tên_khách_hàng\\\":\\\"{user_name}\\\",\\\"địa_chỉ\\\":\\\"{address}\\\",\\\"mã_đồng_hồ\\\":\\\"{device_id}\\\"," +
                                        $"\\\"chỉ_số_nước\\\":\\\"{index_water}\\\",\\\"tháng\\\":\\\"{month}\\\",\\\"năm\\\":\\\"{year}\\\",\\\"số_nhà\\\":\\\"{num_address}\\\",\\\"tên_đường\\\":\\\"{name_street}\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\"}}\"}}]}}";

                                        string generate_text7 = "{\"messages\": [{ \"role\": \"system\", \"content\": " + $"\"{gpt_content_role7}\"" + "}," + assistant_content_role71  + "," + user_content_role7 + assistant_content_role72;
                                        writer.WriteLine(generate_text7);
                                        countFlag7++;
                                    }
                                    
                                }  //comment to check
                                
                            }

                            //============================================================================================================
                            ///*
                            //===============================================SCENARIO SENTENCES==========================================
                             //comment to check
                            gpt_content_role = "You are Bwaco Chatbot. Please interpret the following user input and convert it into JSON of the form " +
                                                "{\\\"hành_động\\\":\\\"\\\",\\\"phân_loại\\\":\\\"\\\",\\\"mã_khách_hàng\\\":\\\"\\\",\\\"số_điện_thoại\\\":\\\"\\\",\\\"tên_khách_hàng\\\":\\\"\\\"," +
                                                "\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\",\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"boolean\\\",\\\"trả_lời\\\":\\\"\\\",\\\"kết_thúc\\\":\\\"boolean\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"\\\"}. Only return JSON. User input:";
                            for (int row = 2; row <= rowCount; row++){
                                string? question = string.IsNullOrEmpty(worksheet.Cells[row, 2]?.Text?.Trim()) ? "" : worksheet.Cells[row, 2]?.Text?.Trim();
                                string? action = string.IsNullOrEmpty(worksheet.Cells[row, 3]?.Text?.Trim()) ? "" : worksheet.Cells[row, 3]?.Text?.Trim();
                                string? category = string.IsNullOrEmpty(worksheet.Cells[row, 4]?.Text?.Trim()) ? "" : worksheet.Cells[row, 4]?.Text?.Trim();
                                string? user_id = string.IsNullOrEmpty(worksheet.Cells[row, 5]?.Text?.Trim()) ? "" : worksheet.Cells[row, 5]?.Text?.Trim();
                                string? phone_number = string.IsNullOrEmpty(worksheet.Cells[row, 6]?.Text?.Trim()) ? "" : worksheet.Cells[row, 6]?.Text?.Trim();
                                string? user_name = string.IsNullOrEmpty(worksheet.Cells[row, 7]?.Text?.Trim()) ? "" : worksheet.Cells[row, 7]?.Text?.Trim();
                                string? address = string.IsNullOrEmpty(worksheet.Cells[row, 8]?.Text?.Trim()) ? "" : worksheet.Cells[row, 8]?.Text?.Trim();
                                string? device_id = string.IsNullOrEmpty(worksheet.Cells[row, 9]?.Text?.Trim()) ? "" : worksheet.Cells[row, 9]?.Text?.Trim();
                                string? index_water = string.IsNullOrEmpty(worksheet.Cells[row, 10]?.Text?.Trim()) ? "" : worksheet.Cells[row, 10]?.Text?.Trim();
                                string? month = string.IsNullOrEmpty(worksheet.Cells[row, 11]?.Text?.Trim()) ? "" : worksheet.Cells[row, 11]?.Text?.Trim();
                                string? year = string.IsNullOrEmpty(worksheet.Cells[row, 12]?.Text?.Trim()) ? "" : worksheet.Cells[row, 12]?.Text?.Trim();
                                string? checkflag = string.IsNullOrEmpty(worksheet.Cells[row, 13]?.Text?.Trim()) ? "" : worksheet.Cells[row, 13]?.Text?.Trim();
                                string? num_address = string.IsNullOrEmpty(worksheet.Cells[row, 14]?.Text?.Trim()) ? "" : worksheet.Cells[row, 14]?.Text?.Trim();
                                string? name_street = string.IsNullOrEmpty(worksheet.Cells[row, 15]?.Text?.Trim()) ? "" : worksheet.Cells[row, 15]?.Text?.Trim();
                                string user_content_role1 = "{\"role\": \"user\", \"content\":" + $"\"{question}\"" + "},";
                                string? ca_nhan = string.IsNullOrEmpty(worksheet.Cells[row, 16]?.Text?.Trim()) ? "" : worksheet.Cells[row, 16]?.Text?.Trim();
                                if(ca_nhan == ""){
                                    ca_nhan = "false";
                                }
                                if(action == "Phản ánh về chất lượng nước" && category == "Nước đục,dơ"){
                                
                                    countKB1 += 1;
                                   
                                    countKB11 = 0;  
                                    for (int row1 = 2; row1 <= rowCount; row1++){
                                        string? question1 = string.IsNullOrEmpty(worksheet.Cells[row1, 2]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 2]?.Text?.Trim();
                                        string? action1 = string.IsNullOrEmpty(worksheet.Cells[row1, 3]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 3]?.Text?.Trim();
                                        string? checkflag1 = string.IsNullOrEmpty(worksheet.Cells[row1, 13]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 13]?.Text?.Trim();
                                        if(action1 == "Trao đổi thông tin" && checkflag1 == "11"){
                                            countKB11 += 1;
                                            if(countKB1 == countKB11){
                                                countKB12 = 0;
                                                for (int row2 = 2; row2 <= rowCount; row2++){
                                                    string? question2 = string.IsNullOrEmpty(worksheet.Cells[row2, 2]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 2]?.Text?.Trim();
                                                    string? action2 = string.IsNullOrEmpty(worksheet.Cells[row2, 3]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 3]?.Text?.Trim();
                                                    string? checkflag2 = string.IsNullOrEmpty(worksheet.Cells[row2, 13]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 13]?.Text?.Trim();
                                                    if(action2 == "Trao đổi thông tin" && checkflag2== "12"){
                                                        countKB12 += 1;
                                                        if(countKB11 == countKB12){
                                                            string assistant_content_role1 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"{action}\\\",\\\"phân_loại\\\":\\\"{category}\\\"," +
                                                            $"\\\"mã_khách_hàng\\\":\\\"{user_id}\\\", \\\"số_điện_thoại\\\":\\\"{phone_number}\\\", \\\"tên_khách_hàng\\\":\\\"{user_name}\\\",\\\"địa_chỉ\\\":\\\"{address}\\\",\\\"mã_đồng_hồ\\\":\\\"{device_id}\\\"," +
                                                            $"\\\"chỉ_số_nước\\\":\\\"{index_water}\\\",\\\"tháng\\\":\\\"{month}\\\",\\\"năm\\\":\\\"{year}\\\",\\\"số_nhà\\\":\\\"{num_address}\\\",\\\"tên_đường\\\":\\\"{name_street}\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Nhà mình dùng nước trực tiếp hay bơm lên bồn ạ?\\\",\\\"kết_thúc\\\":\\\"false\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"1\\\"}}\"}}"; 

                                                            string user_content_role2 = "{\"role\": \"user\", \"content\":" + $"\"{question1}\"" + "},";
                                                            string assistant_content_role2 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"Trao đổi thông tin\\\",\\\"phân_loại\\\":\\\"Thông tin từ khách hàng\\\"," +
                                                            $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                                            $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Tình trạng nước đục xảy ra lâu chưa ạ?\\\",\\\"kết_thúc\\\":\\\"false\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"2\\\"}}\"}}";

                                                            string user_content_role3 = "{\"role\": \"user\", \"content\":" + $"\"{question2}\"" + "},";
                                                            string assistant_content_role3 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"Trao đổi thông tin\\\",\\\"phân_loại\\\":\\\"Thông tin từ khách hàng\\\"," +
                                                            $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                                            $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Thông tin Thông tin của anh/chị đã được ghi nhận. Sẽ có nhân viên của BWACO liên hệ với anh/chị trong thời gian sớm nhất có thể. Xin cảm ơn!\\\",\\\"kết_thúc\\\":\\\"true\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"3\\\"}}\"}}]}}";
                                                            string generate_text = "{\"messages\": [{ \"role\": \"system\", \"content\": " + $"\"{gpt_content_role}\"" + "}," 
                                                                                    + user_content_role1 + assistant_content_role1 + ","
                                                                                    + user_content_role2 + assistant_content_role2 + ","
                                                                                    + user_content_role3 + assistant_content_role3;
                                                            //Console.WriteLine(generate_text);
                                                            // Write the text to a new line
                                                            writer.WriteLine(generate_text);
                                                            break;
                                                        }
                                                        else{
                                                            continue;
                                                        }
                                                    }
                                                }
                                                break;
                                            }
                                            else{
                                                continue;
                                            }
                                                
                                        }

                                    }
                                
                                }
                                
                                
                                else if(action == "Phản ánh về chất lượng nước" && category == "Nước có mùi hôi"){
                                    
                                    countKB5 += 1;
                                    countKB51 = 0;  
                                    for (int row1 = 2; row1 <= rowCount; row1++){
                                        string? question1 = string.IsNullOrEmpty(worksheet.Cells[row1, 2]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 2]?.Text?.Trim();
                                        string? action1 = string.IsNullOrEmpty(worksheet.Cells[row1, 3]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 3]?.Text?.Trim();
                                        string? checkflag1 = string.IsNullOrEmpty(worksheet.Cells[row1, 13]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 13]?.Text?.Trim();
                                        if(action1 == "Trao đổi thông tin" && checkflag1 == "11"){
                                            countKB51 += 1;
                                            
                                            if(countKB5 == countKB51){
                                                countKB52 = 0;
                                                for (int row2 = 2; row2 <= rowCount; row2++){
                                                    string? question2 = string.IsNullOrEmpty(worksheet.Cells[row2, 2]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 2]?.Text?.Trim();
                                                    string? action2 = string.IsNullOrEmpty(worksheet.Cells[row2, 3]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 3]?.Text?.Trim();
                                                    string? checkflag2 = string.IsNullOrEmpty(worksheet.Cells[row2, 13]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 13]?.Text?.Trim();
                                                    if(action2 == "Trao đổi thông tin" && checkflag2== "12"){
                                                        countKB52 += 1;
                                                        if(countKB51 == countKB52){
                                                            string assistant_content_role1 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"{action}\\\",\\\"phân_loại\\\":\\\"{category}\\\"," +
                                                            $"\\\"mã_khách_hàng\\\":\\\"{user_id}\\\", \\\"số_điện_thoại\\\":\\\"{phone_number}\\\", \\\"tên_khách_hàng\\\":\\\"{user_name}\\\",\\\"địa_chỉ\\\":\\\"{address}\\\",\\\"mã_đồng_hồ\\\":\\\"{device_id}\\\"," +
                                                            $"\\\"chỉ_số_nước\\\":\\\"{index_water}\\\",\\\"tháng\\\":\\\"{month}\\\",\\\"năm\\\":\\\"{year}\\\",\\\"số_nhà\\\":\\\"{num_address}\\\",\\\"tên_đường\\\":\\\"{name_street}\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Nhà mình dùng nước trực tiếp hay bơm lên bồn ạ?\\\",\\\"kết_thúc\\\":\\\"false\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"1\\\"}}\"}}"; 

                                                            string user_content_role2 = "{\"role\": \"user\", \"content\":" + $"\"{question1}\"" + "},";
                                                            string assistant_content_role2 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"Trao đổi thông tin\\\",\\\"phân_loại\\\":\\\"Thông tin từ khách hàng\\\"," +
                                                            $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                                            $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Tình trạng nước có mùi hôi xảy ra lâu chưa ạ?\\\",\\\"kết_thúc\\\":\\\"false\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"2\\\"}}\"}}";

                                                            string user_content_role3 = "{\"role\": \"user\", \"content\":" + $"\"{question2}\"" + "},";
                                                            string assistant_content_role3 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"Trao đổi thông tin\\\",\\\"phân_loại\\\":\\\"Thông tin từ khách hàng\\\"," +
                                                            $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                                            $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Thông tin của anh/chị đã được ghi nhận. Sẽ có nhân viên của BWACO liên hệ với anh/chị trong thời gian sớm nhất có thể. Xin cảm ơn!\\\",\\\"kết_thúc\\\":\\\"true\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"3\\\"}}\"}}]}}";
                                                            string generate_text = "{\"messages\": [{ \"role\": \"system\", \"content\": " + $"\"{gpt_content_role}\"" + "}," 
                                                                                    + user_content_role1 + assistant_content_role1 + ","
                                                                                    + user_content_role2 + assistant_content_role2 + ","
                                                                                    + user_content_role3 + assistant_content_role3;
                                                            //Console.WriteLine(generate_text);
                                                            // Write the text to a new line
                                                            writer.WriteLine(generate_text);
                                                            break;
                                                        }
                                                        else{
                                                            continue;
                                                        }
                                                    }
                                                }
                                                break;
                                            }
                                            else{
                                                continue;
                                            }
                                                
                                        }

                                    }
                                
                                }

                                
                                else if(action == "Phản ánh về chất lượng nước" && category == "Nước có bọt trắng,màu trắng đục"){
                                    
                                    countKB6 += 1;
                                    countKB61 = 0;  
                                    for (int row1 = 2; row1 <= rowCount; row1++){
                                        string? question1 = string.IsNullOrEmpty(worksheet.Cells[row1, 2]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 2]?.Text?.Trim();
                                        string? action1 = string.IsNullOrEmpty(worksheet.Cells[row1, 3]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 3]?.Text?.Trim();
                                        string? checkflag1 = string.IsNullOrEmpty(worksheet.Cells[row1, 13]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 13]?.Text?.Trim();
                                        if(action1 == "Trao đổi thông tin" && checkflag1 == "22"){
                                            countKB61 += 1;
                                            if(countKB6 == countKB61){
                                                countKB62 = 0;
                                                for (int row2 = 2; row2 <= rowCount; row2++){
                                                    string? question2 = string.IsNullOrEmpty(worksheet.Cells[row2, 2]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 2]?.Text?.Trim();
                                                    string? action2 = string.IsNullOrEmpty(worksheet.Cells[row2, 3]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 3]?.Text?.Trim();
                                                    string? checkflag2 = string.IsNullOrEmpty(worksheet.Cells[row2, 13]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 13]?.Text?.Trim();
                                                    if(action2 == "Trao đổi thông tin" && checkflag2== "23"){
                                                        countKB62 += 1;
                                                        if(countKB61 == countKB62){
                                                            string assistant_content_role1 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"{action}\\\",\\\"phân_loại\\\":\\\"{category}\\\"," +
                                                            $"\\\"mã_khách_hàng\\\":\\\"{user_id}\\\", \\\"số_điện_thoại\\\":\\\"{phone_number}\\\", \\\"tên_khách_hàng\\\":\\\"{user_name}\\\",\\\"địa_chỉ\\\":\\\"{address}\\\",\\\"mã_đồng_hồ\\\":\\\"{device_id}\\\"," +
                                                            $"\\\"chỉ_số_nước\\\":\\\"{index_water}\\\",\\\"tháng\\\":\\\"{month}\\\",\\\"năm\\\":\\\"{year}\\\",\\\"số_nhà\\\":\\\"{num_address}\\\",\\\"tên_đường\\\":\\\"{name_street}\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Dạ đây chỉ là hiện tượng bọt khí lọt vào đường ống, không ảnh hưởng đến chất lượng nước, anh/chị yên tâm ạ. Xin cảm ơn!\\\",\\\"kết_thúc\\\":\\\"false\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"1\\\"}}\"}}"; 

                                                            string user_content_role2 = "{\"role\": \"user\", \"content\":" + $"\"{question1}\"" + "},";
                                                            string assistant_content_role2 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"Trao đổi thông tin\\\",\\\"phân_loại\\\":\\\"Thông tin từ khách hàng\\\"," +
                                                            $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                                            $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Dạ anh/chị vui lòng xả nước vào 1 ly thủy tinh trong, để một vài phút bọt khí bay lên hết nước sẽ trong trở lại ạ.\\\",\\\"kết_thúc\\\":\\\"false\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"2\\\"}}\"}}";

                                                            string user_content_role3 = "{\"role\": \"user\", \"content\":" + $"\"{question2}\"" + "},";
                                                            string assistant_content_role3 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"Trao đổi thông tin\\\",\\\"phân_loại\\\":\\\"Thông tin từ khách hàng\\\"," +
                                                            $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                                            $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Dạ anh chị cứ theo dõi thêm, cần hỗ trợ vui lòng liên hệ lại với Tổng đài cấp nước ạ. Thông tin của anh/chị đã được ghi nhận. Sẽ có nhân viên của Bwaco liên hệ với anh/chị trong thời gian sớm nhất có thể. Xin cảm ơn!\\\",\\\"kết_thúc\\\":\\\"true\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"3\\\"}}\"}}]}}";
                                                            string generate_text = "{\"messages\": [{ \"role\": \"system\", \"content\": " + $"\"{gpt_content_role}\"" + "}," 
                                                                                    + user_content_role1 + assistant_content_role1 + ","
                                                                                    + user_content_role2 + assistant_content_role2 + ","
                                                                                    + user_content_role3 + assistant_content_role3;
                                                            //Console.WriteLine(generate_text);
                                                            // Write the text to a new line
                                                            writer.WriteLine(generate_text);
                                                            break;
                                                        }
                                                        else{
                                                            continue;
                                                        }
                                                    }
                                                }
                                                break;
                                            }
                                            else{
                                                continue;
                                            }
                                                
                                        }

                                    }
                                    
                                }
                                ///*
                                else if(action == "Yêu cầu sửa chữa đồng hồ"){
                                    
                                    countKB2 += 1;
                                    countKB24 = 0;  
                                    for (int row1 = 2; row1 <= rowCount; row1++){
                                        string? question1 = string.IsNullOrEmpty(worksheet.Cells[row1, 2]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 2]?.Text?.Trim();
                                        string? action1 = string.IsNullOrEmpty(worksheet.Cells[row1, 3]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 3]?.Text?.Trim();
                                        string? category1 = string.IsNullOrEmpty(worksheet.Cells[row1, 4]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 4]?.Text?.Trim();
                                        string? checkflag1 = string.IsNullOrEmpty(worksheet.Cells[row1, 13]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 13]?.Text?.Trim();
                                        if((action1 == "Giải thích,tìm hiểu về chi phí lắp,gắn thêm đồng hồ" && checkflag1 == "14") || (action1 == "Giải thích chung về chi phí" && checkflag1 == "14")){
                                            countKB24 += 1;
                                            if(countKB2 == countKB24){
                                                countKB25 = 0;
                                                for (int row2 = 2; row2 <= rowCount; row2++){
                                                    
                                                    string? question2 = string.IsNullOrEmpty(worksheet.Cells[row2, 2]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 2]?.Text?.Trim();
                                                    string? action2 = string.IsNullOrEmpty(worksheet.Cells[row2, 3]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 3]?.Text?.Trim();
                                                    string? category2 = string.IsNullOrEmpty(worksheet.Cells[row2, 4]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 4]?.Text?.Trim();
                                                    string? checkflag2 = string.IsNullOrEmpty(worksheet.Cells[row2, 13]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 13]?.Text?.Trim();
                                                    if((action2 == "Giải thích,tìm hiểu về chi phí lắp,gắn thêm đồng hồ" && checkflag2 == "15") || (action2 == "Giải thích chung về chi phí" && checkflag2 == "15") ){
                                                        countKB25 += 1;
                                                        if(countKB24 == countKB25){
                                                            countKB26 = 0;
                                                            for (int row3 = 2; row3 <= rowCount; row3++){
                                                                string? question3 = string.IsNullOrEmpty(worksheet.Cells[row3, 2]?.Text?.Trim()) ? "" : worksheet.Cells[row3, 2]?.Text?.Trim();
                                                                string? action3 = string.IsNullOrEmpty(worksheet.Cells[row3, 3]?.Text?.Trim()) ? "" : worksheet.Cells[row3, 3]?.Text?.Trim();
                                                                string? checkflag3 = string.IsNullOrEmpty(worksheet.Cells[row3, 13]?.Text?.Trim()) ? "" : worksheet.Cells[row3, 13]?.Text?.Trim();
                                                                if(action3 == "Trao đổi thông tin" && checkflag3 == "16"){
                                                                    countKB26 += 1;
                                                                    if(countKB25 == countKB26){
                                                                        string assistant_content_role11 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"{action}\\\",\\\"phân_loại\\\":\\\"{category}\\\"," +
                                                                        $"\\\"mã_khách_hàng\\\":\\\"{user_id}\\\", \\\"số_điện_thoại\\\":\\\"{phone_number}\\\", \\\"tên_khách_hàng\\\":\\\"{user_name}\\\",\\\"địa_chỉ\\\":\\\"{address}\\\",\\\"mã_đồng_hồ\\\":\\\"{device_id}\\\"," +
                                                                        $"\\\"chỉ_số_nước\\\":\\\"{index_water}\\\",\\\"tháng\\\":\\\"{month}\\\",\\\"năm\\\":\\\"{year}\\\",\\\"số_nhà\\\":\\\"{num_address}\\\",\\\"tên_đường\\\":\\\"{name_street}\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Thông tin của anh/chị đã được ghi nhận. Sẽ có nhân viên của Bwaco liên hệ với anh/chị trong thời gian sớm nhất có thể.\\\",\\\"kết_thúc\\\":\\\"false\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"1\\\"}}\"}}";

                                                                        string user_content_role21 = "{\"role\": \"user\", \"content\":" + $"\"{question1}\"" + "},";
                                                                        string assistant_content_role21 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"{action1}\\\",\\\"phân_loại\\\":\\\"{category1}\\\"," +
                                                                        $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                                                        $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Trường hợp đồng hồ bị hư, hỏng mất do khách hàng thì sẽ tính phí ạ.\\\",\\\"kết_thúc\\\":\\\"false\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"2\\\"}}\"}}";

                                                                        string user_content_role31 = "{\"role\": \"user\", \"content\":" + $"\"{question2}\"" + "},";
                                                                        string assistant_content_role31 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"{action2}\\\",\\\"phân_loại\\\":\\\"{category2}\\\"," +
                                                                        $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                                                        $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Hiện tại chi phí thay đồng hồ là 500.000đ đến 700.000 đồng tùy thuộc vào vật tư đi kèm ạ.\\\",\\\"kết_thúc\\\":\\\"false\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"3\\\"}}\"}}";

                                                                        string user_content_role41 = "{\"role\": \"user\", \"content\":" + $"\"{question3}\"" + "},";
                                                                        string assistant_content_role41 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"Trao đổi thông tin\\\",\\\"phân_loại\\\":\\\"Thông tin từ khách hàng\\\"," +
                                                                        $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                                                        $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Việc mất/hư hỏng đồng hồ là không ai mong muốn, Công ty đã bàn giao đồng hồ cho khách hàng quản lý nên khi có sự cố phải thay đồng hồ sẽ tính phí. Mong anh/chị thông cảm ạ.\\\",\\\"kết_thúc\\\":\\\"true\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"4\\\"}}\"}}]}}";
                                                                        string generate_text = "{\"messages\": [{ \"role\": \"system\", \"content\": " + $"\"{gpt_content_role}\"" + "}," 
                                                                                                + user_content_role1 + assistant_content_role11 + ","
                                                                                                + user_content_role21 + assistant_content_role21 + ","
                                                                                                + user_content_role31 + assistant_content_role31 + ","
                                                                                                + user_content_role41 + assistant_content_role41;
                                                                        //Console.WriteLine(generate_text);

                                                                        // Write the text to a new line
                                                                        writer.WriteLine(generate_text);            
                                                                    }
                                                                    else{
                                                                        continue;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        else{
                                                            continue;
                                                        }
                                                    }
                                                }
                                            }
                                            else{
                                                continue;
                                            }
                                        }
                                    }
                                    
                                }  
                                
                                else if(action == "Yêu cầu di dời,nâng hạ đồng hồ" && category == "Di dời,chuyển,dịch đồng hồ nước" && checkflag =="789" ){
                                    
                                    countKB3++;
                                    countKB37 = 0;
                                    for (int row1 = 2; row1 <= rowCount; row1++){
                                        string? question1 = string.IsNullOrEmpty(worksheet.Cells[row1, 2]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 2]?.Text?.Trim();
                                        string? action1 = string.IsNullOrEmpty(worksheet.Cells[row1, 3]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 3]?.Text?.Trim();
                                        string? checkflag1 = string.IsNullOrEmpty(worksheet.Cells[row1, 13]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 13]?.Text?.Trim();
                                        if(action1 == "Trao đổi thông tin" && checkflag1 == "17" ){
                                            countKB37 += 1;
                                            if(countKB3 == countKB37){
                                                countKB38 = 0;
                                                for (int row2 = 2; row2 <= rowCount; row2++){
                                                    string? question2 = string.IsNullOrEmpty(worksheet.Cells[row2, 2]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 2]?.Text?.Trim();
                                                    string? action2 = string.IsNullOrEmpty(worksheet.Cells[row2, 3]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 3]?.Text?.Trim();
                                                    string? checkflag2 = string.IsNullOrEmpty(worksheet.Cells[row2, 13]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 13]?.Text?.Trim();
                                                    if(action2 == "Trao đổi thông tin" && checkflag2 == "18"){
                                                        countKB38 += 1;
                                                        if(countKB37 == countKB38){
                                                            countKB39 = 0;
                                                            for (int row3 = 2; row3 <= rowCount; row3++){
                                                                string? question3 = string.IsNullOrEmpty(worksheet.Cells[row3, 2]?.Text?.Trim()) ? "" : worksheet.Cells[row3, 2]?.Text?.Trim();
                                                                string? action3 = string.IsNullOrEmpty(worksheet.Cells[row3, 3]?.Text?.Trim()) ? "" : worksheet.Cells[row3, 3]?.Text?.Trim();
                                                                string? checkflag3 = string.IsNullOrEmpty(worksheet.Cells[row3, 13]?.Text?.Trim()) ? "" : worksheet.Cells[row3, 13]?.Text?.Trim();
                                                                if(action3 == "Trao đổi thông tin" && checkflag3 == "19"){
                                                                    countKB39 += 1;
                                                                    if(countKB38 == countKB39){
                                                                        string assistant_content_role11 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"{action}\\\",\\\"phân_loại\\\":\\\"{category}\\\"," +
                                                                        $"\\\"mã_khách_hàng\\\":\\\"{user_id}\\\", \\\"số_điện_thoại\\\":\\\"{phone_number}\\\", \\\"tên_khách_hàng\\\":\\\"{user_name}\\\",\\\"địa_chỉ\\\":\\\"{address}\\\",\\\"mã_đồng_hồ\\\":\\\"{device_id}\\\"," +
                                                                        $"\\\"chỉ_số_nước\\\":\\\"{index_water}\\\",\\\"tháng\\\":\\\"{month}\\\",\\\"năm\\\":\\\"{year}\\\",\\\"số_nhà\\\":\\\"{num_address}\\\",\\\"tên_đường\\\":\\\"{name_street}\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Anh/chị muốn dời đồng hồ từ trong ra ngoài, hay từ ngoài vào trong ạ?\\\",\\\"kết_thúc\\\":\\\"false\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"1\\\"}}\"}}";

                                                                        string user_content_role21 = "{\"role\": \"user\", \"content\":" + $"\"{question1}\"" + "},";
                                                                        string assistant_content_role21 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"Trao đổi thông tin\\\",\\\"phân_loại\\\":\\\"Thông tin từ khách hàng\\\"," +
                                                                        $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                                                        $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Dạ nếu chỉ dời từ trong ra ngoài, hoặc ngược lại sẽ không mất phí. Trường hợp từ phải qua trái hoặc từ trái qua phải (phát sinh điểm khởi thủy) mới mất phí ạ. Nếu phát sinh chi phí nhân viên Bwaco sẽ lên chiết tính báo trước với anh/chị, nhà mình đồng ý mới làm. Anh/chị cần dời bao xa ạ?\\\",\\\"kết_thúc\\\":\\\"false\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"2\\\"}}\"}}";

                                                                        string user_content_role31 = "{\"role\": \"user\", \"content\":" + $"\"{question2}\"" + "},";
                                                                        string assistant_content_role31 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"Trao đổi thông tin\\\",\\\"phân_loại\\\":\\\"Thông tin từ khách hàng\\\"," +
                                                                        $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                                                        $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Dạ Anh/chị muốn nhân viên xuống dời đồng hồ vào ngày nào ạ?\\\",\\\"kết_thúc\\\":\\\"false\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"3\\\"}}\"}}";

                                                                        string user_content_role41 = "{\"role\": \"user\", \"content\":" + $"\"{question3}\"" + "},";
                                                                        string assistant_content_role41 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"Trao đổi thông tin\\\",\\\"phân_loại\\\":\\\"Thông tin từ khách hàng\\\"," +
                                                                        $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                                                        $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Dạ thời gian Anh/chị yêu cầu đã được ghi nhận và chuyển đến bộ phận kỹ thuật, Bwaco sẽ cố gắng thu xếp để dời sớm cho nhà mình. Trường hợp có các sự cố đột xuất không hỗ trợ kịp mong Anh/chị thông cảm. Xin cảm ơn!\\\",\\\"kết_thúc\\\":\\\"true\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"4\\\"}}\"}}]}}";
                                                                        string generate_text = "{\"messages\": [{ \"role\": \"system\", \"content\": " + $"\"{gpt_content_role}\"" + "}," 
                                                                                                + user_content_role1 + assistant_content_role11 + ","
                                                                                                + user_content_role21 + assistant_content_role21 + ","
                                                                                                + user_content_role31 + assistant_content_role31 + ","
                                                                                                + user_content_role41 + assistant_content_role41;
                                                                        //Console.WriteLine(generate_text);

                                                                        // Write the text to a new line
                                                                        writer.WriteLine(generate_text);            
                                                                    }
                                                                    
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    
                                }
                                
                                else if(action == "Yêu cầu di dời,nâng hạ đồng hồ" && category == "Nâng hạ,chuyển,dời lên xuống đồng hồ nước"){
                                    
                                    countKB7 += 1;
                                   
                                    countKB71 = 0;
                                    for (int row2 = 2; row2 <= rowCount; row2++){
                                        string? question2 = string.IsNullOrEmpty(worksheet.Cells[row2, 2]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 2]?.Text?.Trim();
                                        string? action2 = string.IsNullOrEmpty(worksheet.Cells[row2, 3]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 3]?.Text?.Trim();
                                        string? checkflag2 = string.IsNullOrEmpty(worksheet.Cells[row2, 13]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 13]?.Text?.Trim();
                                        if(action2 == "Trao đổi thông tin" && checkflag2== "26"){
                                            countKB71 += 1;
                                            if(countKB7 == countKB71){
                                                string assistant_content_role1 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"{action}\\\",\\\"phân_loại\\\":\\\"{category}\\\"," +
                                                $"\\\"mã_khách_hàng\\\":\\\"{user_id}\\\", \\\"số_điện_thoại\\\":\\\"{phone_number}\\\", \\\"tên_khách_hàng\\\":\\\"{user_name}\\\",\\\"địa_chỉ\\\":\\\"{address}\\\",\\\"mã_đồng_hồ\\\":\\\"{device_id}\\\"," +
                                                $"\\\"chỉ_số_nước\\\":\\\"{index_water}\\\",\\\"tháng\\\":\\\"{month}\\\",\\\"năm\\\":\\\"{year}\\\",\\\"số_nhà\\\":\\\"{num_address}\\\",\\\"tên_đường\\\":\\\"{name_street}\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Thông tin của anh/chị đã được ghi nhận. Sẽ có nhân viên của Bwaco liên hệ với anh/chị trong vòng 5 ngày làm việc. Dạ Anh/chị muốn nhân viên xuống nâng hạ đồng hồ vào ngày nào ạ?\\\",\\\"kết_thúc\\\":\\\"false\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"1\\\"}}\"}}"; 

                                                string user_content_role2 = "{\"role\": \"user\", \"content\":" + $"\"{question2}\"" + "},";
                                                string assistant_content_role2 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"Trao đổi thông tin\\\",\\\"phân_loại\\\":\\\"Thông tin từ khách hàng\\\"," +
                                                $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                                $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Dạ thời gian Anh/chị yêu cầu đã được ghi nhận và chuyển đến bộ phận kỹ thuật, Bwaco sẽ cố gắng thu xếp để dời sớm cho nhà mình. Trường hợp có các sự cố đột xuất không hỗ trợ kịp mong Anh/chị thông cảm, xin cảm ơn!\\\",\\\"kết_thúc\\\":\\\"true\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"2\\\"}}\"}}]}}";
                                                string generate_text = "{\"messages\": [{ \"role\": \"system\", \"content\": " + $"\"{gpt_content_role}\"" + "}," 
                                                                        + user_content_role1 + assistant_content_role1 + ","
                                                                        + user_content_role2 + assistant_content_role2;
                                                //Console.WriteLine(generate_text);
                                                // Write the text to a new line
                                                writer.WriteLine(generate_text);
                                            }
                                        }
                                    }
                                    
                                }
                                
                                else if(action == "Đăng ký thủ tục lắp,gắn đồng hồ mới"){
                                    countKB4 += 1;
                                    countKB41 = 0;
                                    for (int row1 = 2; row1 <= rowCount; row1++){
                                        string? question1 = string.IsNullOrEmpty(worksheet.Cells[row1, 2]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 2]?.Text?.Trim();
                                        string? action1 = string.IsNullOrEmpty(worksheet.Cells[row1, 3]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 3]?.Text?.Trim();
                                        string? checkflag1 = string.IsNullOrEmpty(worksheet.Cells[row1, 13]?.Text?.Trim()) ? "" : worksheet.Cells[row1, 13]?.Text?.Trim();
                                        if(action1 == "Trao đổi thông tin" && checkflag1 == "20"){
                                            countKB41 += 1;
                                            if(countKB4 == countKB41){
                                                countKB42 = 0;
                                                for (int row2 = 2; row2 <= rowCount; row2++){
                                                    string? question2 = string.IsNullOrEmpty(worksheet.Cells[row2, 2]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 2]?.Text?.Trim();
                                                    string? action2 = string.IsNullOrEmpty(worksheet.Cells[row2, 3]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 3]?.Text?.Trim();
                                                    string? checkflag2 = string.IsNullOrEmpty(worksheet.Cells[row2, 13]?.Text?.Trim()) ? "" : worksheet.Cells[row2, 13]?.Text?.Trim();
                                                    if(action2 == "Trao đổi thông tin" && checkflag2== "21"){
                                                        countKB42 += 1;
                                                        if(countKB42 == countKB41){
                                                            string assistant_content_role11 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"{action}\\\",\\\"phân_loại\\\":\\\"{category}\\\"," +
                                                            $"\\\"mã_khách_hàng\\\":\\\"{user_id}\\\", \\\"số_điện_thoại\\\":\\\"{phone_number}\\\", \\\"tên_khách_hàng\\\":\\\"{user_name}\\\",\\\"địa_chỉ\\\":\\\"{address}\\\",\\\"mã_đồng_hồ\\\":\\\"{device_id}\\\"," +
                                                            $"\\\"chỉ_số_nước\\\":\\\"{index_water}\\\",\\\"tháng\\\":\\\"{month}\\\",\\\"năm\\\":\\\"{year}\\\",\\\"số_nhà\\\":\\\"{num_address}\\\",\\\"tên_đường\\\":\\\"{name_street}\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Dạ, hồ sơ lắp mới bao gồm: bản photo Giấy chủ quyền nhà/đất, căn cước công dân của chủ hộ. Anh/chị vui lòng cung cấp địa chỉ lắp đặt đồng hồ, tên chủ hộ, số điện thoại ạ.\\\",\\\"kết_thúc\\\":\\\"false\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"1\\\"}}\"}}";

                                                            string user_content_role21 = "{\"role\": \"user\", \"content\":\"Tên khách hàng: Nguyễn A, số điện thoại:  0123456789, địa chỉ: 03 Tô Ngọc Vân\"},";
                                                            string assistant_content_role21 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"Cung cấp thông tin khách hàng\\\",\\\"phân_loại\\\":\\\"Thông tin khách hàng\\\"," +
                                                            $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"0123456789\\\", \\\"tên_khách_hàng\\\":\\\"Nguyễn A\\\",\\\"địa_chỉ\\\":\\\"03 Tô Ngọc Vân\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                                            $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"03\\\",\\\"tên_đường\\\":\\\"Tô Ngọc Vân\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Anh/chị đã xây nhà rồi hay đất trống ạ?\\\",\\\"kết_thúc\\\":\\\"false\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"2\\\"}}\"}}";

                                                            string user_content_role31 = "{\"role\": \"user\", \"content\":" + $"\"{question1}\"" + "},";
                                                            string assistant_content_role31 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"Trao đổi thông tin\\\",\\\"phân_loại\\\":\\\"Thông tin từ khách hàng\\\"," +
                                                            $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                                            $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Trường hợp chưa xây nhà Anh/chị vui lòng chuẩn bị hồ sơ bao gồm: bản photo Giấy chủ quyền nhà/đất, Căn cước công dân của chủ hộ và Giấy phép xây dựng.\\\",\\\"kết_thúc\\\":\\\"false\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"3\\\"}}\"}}";

                                                            string user_content_role41 = "{\"role\": \"user\", \"content\":" + $"\"{question2}\"" + "},";
                                                            string assistant_content_role41 = $"{{ \"role\": \"assistant\", \"content\": \"{{\\\"hành_động\\\":\\\"Trao đổi thông tin\\\",\\\"phân_loại\\\":\\\"Thông tin từ khách hàng\\\"," +
                                                            $"\\\"mã_khách_hàng\\\":\\\"\\\", \\\"số_điện_thoại\\\":\\\"\\\", \\\"tên_khách_hàng\\\":\\\"\\\",\\\"địa_chỉ\\\":\\\"\\\",\\\"mã_đồng_hồ\\\":\\\"\\\"," +
                                                            $"\\\"chỉ_số_nước\\\":\\\"\\\",\\\"tháng\\\":\\\"\\\",\\\"năm\\\":\\\"\\\",\\\"số_nhà\\\":\\\"\\\",\\\"tên_đường\\\":\\\"\\\",\\\"cơ_quan_doanh_nghiệp\\\":\\\"false\\\",\\\"trả_lời\\\":\\\"Dạ trường hợp đất trống phải bổ sung giấy phép xây dựng, nếu chưa có nhà mình có thể bổ sung sau ạ. Xin cảm ơn!\\\",\\\"kết_thúc\\\":\\\"true\\\",\\\"số_thứ_tự_trả_lời\\\":\\\"4\\\"}}\"}}]}}";

                                                            
                                                            string generate_text = "{\"messages\": [{ \"role\": \"system\", \"content\": " + $"\"{gpt_content_role}\"" + "}," 
                                                                                    + user_content_role1 + assistant_content_role11 + ","
                                                                                    + user_content_role21 + assistant_content_role21 + ","
                                                                                    + user_content_role31 + assistant_content_role31 + ","
                                                                                    + user_content_role41 + assistant_content_role41;
                                                            //Console.WriteLine(generate_text);

                                                            // Write the text to a new line
                                                            writer.WriteLine(generate_text);       
                                                        }                      
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            } //*/       //comment to check                      
                        }         
                    }     
                }

                Console.WriteLine("Create file successfully!!!");
                Console.WriteLine("Count Single: " + countSingle);
                Console.WriteLine("Count F0: " + countFlag0);
                Console.WriteLine("Count F1: " + countFlag1);
                Console.WriteLine("Count F2: " + countFlag2);
                Console.WriteLine("Count F3: " + countFlag3);
                Console.WriteLine("Count F4: " + countFlag4);
                Console.WriteLine("Count F5: " + countFlag5);
                Console.WriteLine("Count F6: " + countFlag6);
                Console.WriteLine("Count F7: " + countFlag7);
                Console.WriteLine("Count KB1: " + countKB1);
                Console.WriteLine("Count KB2: " + countKB2);
                Console.WriteLine("Count KB3: " + countKB3);
                Console.WriteLine("Count KB4: " + countKB4);
                Console.WriteLine("Count KB5: " + countKB5);
                Console.WriteLine("Count KB6: " + countKB6);
                Console.WriteLine("Count KB7: " + countKB7);

                Console.WriteLine("Count Total Flag: " + (countFlag0 + countFlag1 + countFlag2 + countFlag3 + 
                                                            countFlag4 + countFlag5 + countFlag6 + countFlag7));
                Console.WriteLine("Count Total All: " + (countFlag0 + countFlag1 + countFlag2 + countFlag3 + 
                                                        countFlag4 + countFlag5 + countFlag6 + countFlag7 + countSingle +
                                                        countKB1 + countKB2 + countKB3 + countKB4 + countKB5 + countKB6 + countKB7));
            }
            catch (Exception ex){
                Console.WriteLine("Error: " + ex.Message);
            }
    }
        private static void Main(string[] args){
            make_file_Click();
            return;
    }

}
}