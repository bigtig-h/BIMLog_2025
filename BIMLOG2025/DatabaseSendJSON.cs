using MySql.Data.MySqlClient;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace BIMLOG2025
{
    public class DatabaseSendJSON
    {
        public static string receivedFolderPath;
        public static string user_Name;
        public static JObject dataFromLogger;
        public static string server = "10.35.119.137";
        public static string user = "data_user";
        public static string password = "3846";
        public static string database = "bimloggerdb";
        public static int port = 3306;
        public static void OnDataReceived( Object? sender, JObject data )
        {
            dataFromLogger = data;
            string connectionString = $"server={server};port={port};user={user};password={password};";

            JObject bimLog = (JObject)dataFromLogger["bimlog"];
            string userName = bimLog?["userName"]?.ToString();
            string projectGUID = bimLog?["projectGUID"]?.ToString();
            string projectName = bimLog?["projectName"]?.ToString();
            string startTime = bimLog?["startTime"]?.ToString();
            string endTime = bimLog?["endTime"]?.ToString();
            var checkSaved = bimLog?["Saved"]?.ToString();
            bool saved = false;
            if (checkSaved == "True")
            {
                saved = true;
            }

            try
            {
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                {
                    connection.Open();

                    // 데이터베이스 존재 여부 확인 및 생성
                    string createDatabaseQuery = $"CREATE DATABASE IF NOT EXISTS {database};";
                    using (MySqlCommand cmd = new MySqlCommand(createDatabaseQuery, connection))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    // 데이터베이스 선택
                    connection.ChangeDatabase(database);
                    byte[] jsonData = Encoding.UTF8.GetBytes(dataFromLogger.ToString());
                    string saveData = @"INSERT INTO user_file (User_ID, File_GUID, File_Name, start_time, end_time, File_Saved, File) VALUES ((SELECT User_ID FROM user WHERE revit_id = @userName), @projectGUID, @projectName, @startTime, @endTime, @saved, @file);";
                    using (MySqlCommand cmd = new MySqlCommand(saveData, connection))
                    {
                        cmd.Parameters.Add("@userName", MySqlDbType.VarChar).Value = userName;
                        cmd.Parameters.Add("@projectGUID", MySqlDbType.VarChar).Value = projectGUID;
                        cmd.Parameters.Add("@projectName", MySqlDbType.VarChar).Value = projectName;
                        cmd.Parameters.Add("@startTime", MySqlDbType.DateTime).Value = startTime;
                        cmd.Parameters.Add("@endTime", MySqlDbType.DateTime).Value = endTime;
                        cmd.Parameters.Add("@saved", MySqlDbType.Bit).Value = saved;
                        cmd.Parameters.Add("@file", MySqlDbType.LongBlob).Value = jsonData;
                        cmd.ExecuteNonQuery();
                    }
                    Debug.WriteLine("Data inserted Successfully.");

                }
            }
            catch (MySqlException ex)
            {
                Debug.WriteLine("mysql Error occured: " + ex);
            }

        }

        public static void BackUpDataReceived( Object? sender, string path )
        {
            string connectionString = $"server={server};port={port};user={user};password={password};";
            receivedFolderPath = path;
            var fileArray = Directory.GetFiles(receivedFolderPath, "*.json");
            var fileNum = fileArray.Length;
            DateTime currentDate = DateTime.Now;
            DateTime comparisonDate = currentDate.AddDays(-7);
            string dateFormat = "yyyy-MM-dd_HH-mm-ss";
            List<string> oldFiles = GetOldFiles(receivedFolderPath, currentDate);
            List<(string GUID, string FileName, string sTime, bool Saved)> dbFiles = GetDataFromDatabase(connectionString, comparisonDate);
            List<string> filesToUpdate = new List<string>();

            foreach (string file in oldFiles)
            {
                string fileName = Path.GetFileNameWithoutExtension(file);
                string[] parts = fileName.Split('_');
                string file_date = parts[0] + "_" + parts[1];
                string file_guid = parts[2];
                string file_Name = string.Join("_", parts, 3, parts.Length - 3);
                bool isSaved = false;
                if (file_Name.EndsWith("_saved"))
                {
                    isSaved = true;
                    file_Name = file_Name.Substring(0, file_Name.Length - 6);
                }
                bool existInDb = false;
                foreach (var dbFile in dbFiles)
                {
                    var time = DateTime.Parse(dbFile.sTime).ToString(dateFormat);
                    if (time == file_date && dbFile.GUID == file_guid && dbFile.FileName == file_Name && dbFile.Saved == isSaved)
                    {
                        existInDb = true;
                    }
                }
                if (!existInDb)
                {
                    filesToUpdate.Add(file);
                }
            }

            if (filesToUpdate.Count != 0)
            {
                try
                {
                    // 여기에 이제 filesToUpdate 에서 파일 읽어와서 Insert 문 작성
                    foreach (string file in filesToUpdate)
                    {
                        string jsonContent = File.ReadAllText(file);
                        // 여기서 err
                        JObject jsonObj = JObject.Parse(jsonContent);
                        JObject bimLog = jsonObj["bimlog"] as JObject;

                        string userName = bimLog?["userName"]?.ToString();
                        string projectGUID = bimLog?["projectGUID"]?.ToString();
                        string projectName = bimLog?["projectName"]?.ToString();
                        string startTime = bimLog?["startTime"]?.ToString();
                        string endTime = bimLog?["endTime"]?.ToString();
                        var checkSaved = bimLog?["Saved"]?.ToString();
                        bool saved = false;
                        if (checkSaved == "True")
                        {
                            saved = true;
                        }

                        using (MySqlConnection connection = new MySqlConnection(connectionString))
                        {
                            connection.Open();
                            connection.ChangeDatabase(database);
                            byte[] jsonData = Encoding.UTF8.GetBytes(jsonContent);
                            string saveData = $@"INSERT INTO {database}.user_file
(User_ID, File_GUID, File_Name, start_time, end_time, File_Saved, File) VALUES
((SELECT User_ID FROM {database}.user WHERE revit_id = @userName)
, @projectGUID, @projectName, @startTime, @endTime, @saved, @file);";
                            using (MySqlCommand cmd = new MySqlCommand(saveData, connection))
                            {
                                cmd.Parameters.Add("@userName", MySqlDbType.VarChar).Value = userName;
                                cmd.Parameters.Add("@projectGUID", MySqlDbType.VarChar).Value = projectGUID;
                                cmd.Parameters.Add("@projectName", MySqlDbType.VarChar).Value = projectName;
                                cmd.Parameters.Add("@startTime", MySqlDbType.DateTime).Value = startTime;
                                cmd.Parameters.Add("@endTime", MySqlDbType.DateTime).Value = endTime;
                                cmd.Parameters.Add("@saved", MySqlDbType.Bit).Value = saved;
                                cmd.Parameters.Add("@file", MySqlDbType.LongBlob).Value = jsonData;
                                cmd.ExecuteNonQuery();
                            }
                            Debug.WriteLine("Data inserted Successfully.");
                        }
                    }
                }
                catch (MySqlException myex)
                {
                    Debug.WriteLine("mysql Error occured: " + myex);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Error occured: " + ex);
                }
            }
        }
        static List<string> GetOldFiles( string folderPath, DateTime thresholdDate )
        {
            List<string> oldFiles = new List<string>();
            if (Directory.GetFiles(folderPath, "*.json").Length != 0)
            {
                var getUserFile = Directory.GetFiles(folderPath, "*.json").GetValue(0) as string;
                string x = File.ReadAllText(getUserFile);
                JObject bimlog = JObject.Parse(x);
                bimlog = bimlog["bimlog"] as JObject;
                user_Name = bimlog?["userName"]?.ToString();
            }

            foreach (string file in Directory.GetFiles(folderPath, "*.json"))
            {
                string[] fileParts = Path.GetFileNameWithoutExtension(file).Split('_');
                DateTime fileDate = DateTime.ParseExact(fileParts[0] + "_" + fileParts[1], "yyyy-MM-dd_HH-mm-ss", null);

                if (fileDate < thresholdDate && fileDate > thresholdDate.AddDays(-7))
                {
                    oldFiles.Add(file);
                }
            }
            return oldFiles;
        }
        static List<(string GUID, string FileName, string sTime, bool Saved)> GetDataFromDatabase( string connectionString, DateTime sevenDaysAgo )
        {
            List<(string GUID, string FileName, string sTime, bool Saved)> data = new List<(string GUID, string FileName, string sTime, bool Saved)>();

            if (user_Name != null)
            {

                using (MySqlConnection conn = new MySqlConnection(connectionString))
                {
                    conn.Open();
                    string query = $@"
                SELECT File_GUID, File_Name, start_time, File_Saved
                FROM {database}.user_file
                WHERE User_ID = (SELECT User_ID FROM {database}.user WHERE revit_id = @revit_ID)  AND start_time >= @SevenDaysAgo";

                    MySqlCommand cmd = new MySqlCommand(query, conn);
                    cmd.Parameters.AddWithValue("@SevenDaysAgo", sevenDaysAgo);
                    cmd.Parameters.Add("@revit_ID", MySqlDbType.VarChar).Value = user_Name;

                    using (MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            reader["File_Saved"].ToString();
                            data.Add((
                                reader["File_GUID"].ToString(),
                                reader["File_Name"].ToString(),
                                reader["start_time"].ToString(),
                                reader.GetInt32(reader.GetOrdinal("File_Saved")) == 1
                                ));

                        }
                    }
                }
            }

            return data;
        }

    }
}
