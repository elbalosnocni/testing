using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using zkemkeeper;

namespace RonaldJackFlex
{
    class MachineInfo
    {
        public string Name { get; set; }
        public int Id { get; set; }
        public string IP { get; set; }
        public int Port { get; set; }
    }

    class AttendanceRecord
    {
        public string MachineName { get; set; }
        public string UserID { get; set; }
        public string DateTimeCheck { get; set; }
        public string CheckType { get; set; }
    }

    class Program
    {
        static string GoogleScriptUrl = "";
        static List<MachineInfo> machineList = new List<MachineInfo>();

        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            
            if (!LoadConfig()) { 
                Console.WriteLine("Nhấn phím bất kỳ để thoát...");
                Console.ReadKey(); 
                return; 
            }

            ConcurrentBag<AttendanceRecord> allRecords = new ConcurrentBag<AttendanceRecord>();
            Console.WriteLine("--- BẮT ĐẦU KẾT NỐI ĐA LUỒNG TỚI CÁC MÁY CHẤM CÔNG ---");

            Parallel.ForEach(machineList, machine =>
            {
                CZKEM axCZKEM = new CZKEM();
                if (axCZKEM.Connect_Net(machine.IP, machine.Port))
                {
                    Console.WriteLine(string.Format("[{0}] Kết nối thành công tới IP {1}.", machine.Name, machine.IP));
                    if (axCZKEM.ReadGeneralLogData(machine.Id))
                    {
                        string enrollNumber = "";
                        int verifyMode = 0, inOutMode = 0;
                        int year = 0, month = 0, day = 0, hour = 0, minute = 0, second = 0, workCode = 0;

                        while (axCZKEM.SSR_GetGeneralLogData(machine.Id, out enrollNumber, out verifyMode, 
                               out inOutMode, out year, out month, out day, out hour, out minute, out second, ref workCode))
                        {
                            string fullTime = string.Format("{0:D2}/{1:D2}/{2} {3:D2}:{4:D2}:{5:D2}", day, month, year, hour, minute, second);
                            string type = (inOutMode == 0) ? "Vào" : (inOutMode == 1 ? "Ra" : "Khác");

                            allRecords.Add(new AttendanceRecord
                            {
                                MachineName = machine.Name,
                                UserID = enrollNumber,
                                DateTimeCheck = fullTime,
                                CheckType = type
                            });
                        }
                    }
                    axCZKEM.Disconnect();
                    Console.WriteLine(string.Format("[{0}] Đã hoàn tất tải dữ liệu và ngắt kết nối.", machine.Name));
                }
                else
                {
                    Console.WriteLine(string.Format("[LỖI] Không thể kết nối tới máy [{0}] tại địa chỉ IP: {1}", machine.Name, machine.IP));
                }
            });

            if (allRecords.Count > 0)
            {
                Console.WriteLine(string.Format("\nĐang tiến hành đồng bộ {0} dữ liệu lên Google Sheets...", allRecords.Count));
                SendToGoogle(allRecords);
            }
            else
            {
                Console.WriteLine("\nKhông quét được dữ liệu chấm công mới nào từ các thiết bị.");
            }

            Console.WriteLine("\n=== TOÀN BỘ TIẾN TRÌNH HOÀN TẤT ===");
            Console.WriteLine("Nhấn phím bất kỳ để đóng cửa sổ...");
            Console.ReadKey();
        }

        static bool LoadConfig()
        {
            string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.txt");
            if (!File.Exists(configPath))
            {
                Console.WriteLine("[LỖI] Không tìm thấy tệp cấu hình 'config.txt' nằm chung thư mục!");
                return false;
            }

            try
            {
                foreach (var line in File.ReadAllLines(configPath))
                {
                    string trimmed = line.Trim();
                    if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith("#")) continue;

                    if (trimmed.StartsWith("URL="))
                    {
                        GoogleScriptUrl = trimmed.Substring(4).Trim();
                    }
                    else if (trimmed.StartsWith("MÁY="))
                    {
                        var parts = trimmed.Substring(4).Split('|');
                        if (parts.Length == 4)
                        {
                            machineList.Add(new MachineInfo
                            {
                                Name = parts[0].Trim(),
                                Id = int.Parse(parts[1].Trim()),
                                IP = parts[2].Trim(),
                                Port = int.Parse(parts[3].Trim())
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("[LỖI] Định dạng dữ liệu trong file config.txt không hợp lệ: {0}", ex.Message));
                return false;
            }

            if (string.IsNullOrEmpty(GoogleScriptUrl))
            {
                Console.WriteLine("[LỖI] Chưa cấu hình đường dẫn URL của Google Web App trong config.txt!");
                return false;
            }
            if (machineList.Count == 0)
            {
                Console.WriteLine("[LỖI] Danh sách máy chấm công trong config.txt đang trống!");
                return false;
            }

            return true;
        }

        // Hàm hỗ trợ mã hóa chuỗi JSON an toàn, loại bỏ ký tự điều khiển lỗi
        static string EscapeJson(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            StringBuilder sb = new StringBuilder();
            foreach (char c in s)
            {
                switch (c)
                {
                    case '\"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\b': sb.Append("\\b"); break;
                    case '\f': sb.Append("\\f"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < ' ') {
                            // Bỏ qua hoặc biến thành khoảng trắng đối với ký tự điều khiển không hợp lệ
                            sb.Append(" ");
                        } else {
                            sb.Append(c);
                        }
                        break;
                }
            }
            return sb.ToString();
        }

        static void SendToGoogle(ConcurrentBag<AttendanceRecord> records)
        {
            using (WebClient client = new WebClient())
            {
                try
                {
                    client.Encoding = Encoding.UTF8;
                    client.Headers[HttpRequestHeader.ContentType] = "application/json";

                    StringBuilder json = new StringBuilder("[");
                    foreach (var r in records)
                    {
                        json.Append("{");
                        json.Append("\"MachineName\":\"" + EscapeJson(r.MachineName) + "\",");
                        json.Append("\"UserID\":\"" + EscapeJson(r.UserID) + "\",");
                        json.Append("\"DateTimeCheck\":\"" + EscapeJson(r.DateTimeCheck) + "\",");
                        json.Append("\"CheckType\":\"" + EscapeJson(r.CheckType) + "\"");
                        json.Append("},");
                    }
                    if (json.Length > 1) json.Length--;
                    json.Append("]");

                    string response = client.UploadString(GoogleScriptUrl, "POST", json.ToString());
                    Console.WriteLine(string.Format("[THÀNH CÔNG] Đồng bộ hoàn tất! Phản hồi từ Google: {0}", response));
                }
                catch (Exception ex)
                {
                    Console.WriteLine(string.Format("[LỖI HỆ THỐNG] Trục trặc đường truyền tới Google: {0}", ex.Message));
                }
            }
        }
    }
}
