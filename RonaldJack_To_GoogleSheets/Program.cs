using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
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
        static string GoogleScriptUrl = "https://script.google.com/macros/s/AKfycbz_qDgGg-iEfVDmyzf-YzNrr8720YsYDXorpD7xP_C8ptGB-GQchpzI4tXelr86VMhaZA/exec";
        static List<MachineInfo> machineList = new List<MachineInfo>();

        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            if (!LoadConfig()) { Console.ReadKey(); return; }

            ConcurrentBag<AttendanceRecord> allRecords = new ConcurrentBag<AttendanceRecord>();
            Console.WriteLine("--- BẮT ĐẦU KẾT NỐI ĐA LUỒNG ---");

            Parallel.ForEach(machineList, machine =>
            {
                CZKEM axCZKEM = new CZKEM();
                if (axCZKEM.Connect_Net(machine.IP, machine.Port))
                {
                    Console.WriteLine($"[{machine.Name}] Kết nối thành công.");
                    if (axCZKEM.ReadGeneralLogData(machine.Id))
                    {
                        string enrollNumber = "";
                        int verifyMode = 0, inOutMode = 0;
                        int year = 0, month = 0, day = 0, hour = 0, minute = 0, second = 0, workCode = 0;

                        while (axCZKEM.SSR_GetGeneralLogData(machine.Id, out enrollNumber, out verifyMode, 
                               out inOutMode, out year, out month, out day, out hour, out minute, out second, ref workCode))
                        {
                            allRecords.Add(new AttendanceRecord
                            {
                                MachineName = machine.Name,
                                UserID = enrollNumber,
                                DateTimeCheck = $"{day:D2}/{month:D2}/{year} {hour:D2}:{minute:D2}:{second:D2}",
                                CheckType = (inOutMode == 0) ? "Vào" : (inOutMode == 1 ? "Ra" : "Khác")
                            });
                        }
                    }
                    axCZKEM.Disconnect();
                }
                else
                {
                    Console.WriteLine($"[LỖI] Thất bại khi kết nối tới {machine.Name} ({machine.IP})");
                }
            });

            if (allRecords.Count > 0)
            {
                Console.WriteLine($"\nĐang đẩy {allRecords.Count} dòng lên Google Sheets...");
                Task.Run(() => SendToGoogle(allRecords)).Wait();
            }
            else
            {
                Console.WriteLine("Không có dữ liệu mới.");
            }

            Console.WriteLine("\nHoàn tất tiến trình. Nhấn phím bất kỳ để đóng...");
            Console.ReadKey();
        }

        static bool LoadConfig()
        {
            string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.txt");
            if (!File.Exists(configPath))
            {
                Console.WriteLine("Lỗi: Không tìm thấy file config.txt!");
                return false;
            }

            foreach (var line in File.ReadAllLines(configPath))
            {
                string trimmed = line.Trim();
                if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith("#")) continue;

                if (trimmed.StartsWith("URL="))
                {
                    GoogleScriptUrl = trimmed.Substring(4);
                }
                else if (trimmed.StartsWith("MÁY="))
                {
                    var parts = trimmed.Substring(4).Split('|');
                    if (parts.Length == 4)
                    {
                        machineList.Add(new MachineInfo {
                            Name = parts[0], Id = int.Parse(parts[1]), IP = parts[2], Port = int.Parse(parts[3])
                        });
                    }
                }
            }
            return !string.IsNullOrEmpty(GoogleScriptUrl) && machineList.Count > 0;
        }

        static async Task SendToGoogle(ConcurrentBag<AttendanceRecord> records)
        {
            using (HttpClient client = new HttpClient())
            {
                try
                {
                    // Tự dựng chuỗi JSON đơn giản để không cần cài thêm thư viện phụ trợ bên ngoài
                    StringBuilder json = new StringBuilder("[");
                    foreach (var r in records)
                    {
                        json.Append($"{{\"MachineName\":\"{r.MachineName}\",\"UserID\":\"{r.UserID}\",\"DateTimeCheck\":\"{r.DateTimeCheck}\",\"CheckType\":\"{r.CheckType}\"}},");
                    }
                    if (json.Length > 1) json.Length--; // Bỏ dấu phẩy cuối
                    json.Append("]");

                    var content = new StringContent(json.ToString(), Encoding.UTF8, "application/json");
                    HttpResponseMessage response = await client.PostAsync(GoogleScriptUrl, content);
                    if (response.IsSuccessStatusCode)
                        Console.WriteLine("[THÀNH CÔNG] Dữ liệu đã hiển thị trên Google Sheet!");
                    else
                        Console.WriteLine($"[LỖI] HTTP Status: {response.StatusCode}");
                }
                catch (Exception ex) { Console.WriteLine($"[LỖI HỆ THỐNG]: {ex.Message}"); }
            }
        }
    }
}
