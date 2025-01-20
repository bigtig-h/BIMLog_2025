#region Namespaces
using Autodesk.Internal.Windows;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.UI.Selection;
using Autodesk.Windows;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using MySql.Data;
using MySql;
#endregion 


namespace BIMLOG2025
{

    [Transaction(TransactionMode.Manual)]


    public class JLogger : IExternalApplication
    {
        public int checkRunNum;
        public static JLogger thisApp;
        public string folderPath;
        public string JsonFile;
        public string userId;
        public string fileName;
        // 24.5.14. 수정
        public Dictionary<string, JObject> fileAndJObject = new Dictionary<string, JObject>();
        public Dictionary<string, string> fileAndPath = new Dictionary<string, string>();
        public int newFileCount;
        // JSON FIELD
        public string beginLog = @"{ 'bimlog' : { } }";
        // 24.5.22. 수정
        public List<ElementId> addedElementList = new List<ElementId>();
        public List<ElementId> selectedElementList = new List<ElementId>();
        public List<ElementId> addedStairList = new List<ElementId>();
        // jobject 이놈이 계속해서 쓰이는 JObject
        public JObject publicJobject = new JObject();
        // Log를 담을 jarray : 계속해서 로그 누적시킴.
        public JArray jarray = new JArray();

        // 24.7.4. <<<mySQL 연동>>>
        // 데이터 전송을 위한 이벤트 
        public event EventHandler<JObject> DataSent;
        public event EventHandler<string> BackUpData;
        // 이벤트를 발생시키는 함수
        protected virtual void SendData( JObject data )
        {
            DataSent?.Invoke(this, data);
        }

        protected virtual void BackUp( string folderPath )
        {
            BackUpData?.Invoke(this, folderPath);
        }

        #region EventHandlers


        //Log file path designation
        private void SetLogPath( )
        {
            try
            {
                FileInfo fi = new FileInfo("C:\\ProgramData\\Autodesk\\Revit\\BIG_Log Directory2.txt");
                if (fi.Exists) //if Folder Path 가 존재한다면
                {
                    string LogFilePath = "C:\\ProgramData\\Autodesk\\Revit";
                    string pathFile = Path.Combine(LogFilePath, "BIG_Log Directory2.txt");
                    using (StreamReader readtext = new StreamReader(pathFile, true))
                    {
                        string readText = readtext.ReadLine();
                        Console.WriteLine(readText);
                        folderPath = readText;
                    }
                }
                else
                {
                    FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
                    folderBrowser.Description = "Select a folder to save Revit Modeling Log file.";
                    folderBrowser.ShowNewFolderButton = true;
                    // folderBrowser.SelectedPath = Folder_Path;
                    //folderBrowser.RootFolder = Environment.SpecialFolder.Personal;
                    //folderBrowser.SelectedPath = project_name.Properties.Settings.Default.Folder_Path;
                    if (folderBrowser.ShowDialog() == DialogResult.OK)
                    {
                        folderPath = folderBrowser.SelectedPath;
                        string LogFilePath = "C:\\ProgramData\\Autodesk\\Revit";
                        string pathFile = Path.Combine(LogFilePath, "BIG_Log Directory2.txt");
                        using (StreamWriter writetext = new StreamWriter(pathFile, true))
                        {
                            writetext.WriteLine(folderPath);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Autodesk.Revit.UI.TaskDialog.Show("ADD-IN FAILED", ex.Message);
            }
            //JsonFile = Path.Combine(folderPath, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss") + ".json");
            //Debug.WriteLine(JsonFile);
        }


        public Result OnStartup( UIControlledApplication application )
        {
            thisApp = this;
            try
            {
                if (string.IsNullOrEmpty(folderPath))
                {
                    SetLogPath();
                }
                ComponentManager.ItemExecuted += new EventHandler<RibbonItemExecutedEventArgs>(CommandExecuted);
                application.ControlledApplication.DocumentChanged += new EventHandler<DocumentChangedEventArgs>(DocumentChangeTracker);
                application.ControlledApplication.FailuresProcessing += new EventHandler<FailuresProcessingEventArgs>(FailureTracker);
                application.ControlledApplication.DocumentOpened += new EventHandler<DocumentOpenedEventArgs>(DocumentOpenedTracker);
                application.ControlledApplication.DocumentCreated += new EventHandler<DocumentCreatedEventArgs>(DocumentCreatedTracker);
                application.ControlledApplication.DocumentClosing += new EventHandler<DocumentClosingEventArgs>(DocumentClosingTracker);
                application.ControlledApplication.DocumentSavedAs += new EventHandler<DocumentSavedAsEventArgs>(DocumentSavedAsTracker);
                application.ControlledApplication.DocumentSaving += new EventHandler<DocumentSavingEventArgs>(DocumentSavingTracker);
                application.SelectionChanged += new EventHandler<SelectionChangedEventArgs>(SelectionChangeTracker);
                // JSON 
                JLogger.thisApp.DataSent += DatabaseSendJSON.OnDataReceived;
                JLogger.thisApp.BackUpData += DatabaseSendJSON.BackUpDataReceived;


                BackUp(folderPath);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("err: " + ex);
                return Result.Failed;
            }
            return Result.Succeeded;
        }
        public Result OnShutdown( UIControlledApplication application )
        {
            try
            {
                ComponentManager.ItemExecuted -= CommandExecuted;
                application.ControlledApplication.DocumentChanged -= DocumentChangeTracker;
                application.ControlledApplication.FailuresProcessing -= FailureTracker;
                application.ControlledApplication.DocumentOpened -= DocumentOpenedTracker;
                application.ControlledApplication.DocumentCreated -= DocumentCreatedTracker;
                application.ControlledApplication.DocumentClosing -= DocumentClosingTracker;
                application.ControlledApplication.DocumentSavedAs -= DocumentSavedAsTracker;
                application.ControlledApplication.DocumentSaving -= DocumentSavingTracker;
                application.SelectionChanged -= SelectionChangeTracker;
                JLogger.thisApp.DataSent -= DatabaseSendJSON.OnDataReceived;
                JLogger.thisApp.BackUpData -= DatabaseSendJSON.BackUpDataReceived;

            }
            catch (Exception)
            {
                return Result.Failed;
            }
            return Result.Succeeded;
        }
        // Doc 열렸을 때 동작하는 Tracker
        public void DocumentOpenedTracker( object sender, DocumentOpenedEventArgs e )
        {
            Document doc = e.Document;
            userId = doc.Application.Username;
            string filename = doc.PathName;
            // 프로젝트 GUID
            BasicFileInfo info = BasicFileInfo.Extract(filename);
            DocumentVersion v = info.GetDocumentVersion();
            string projectGUID = v.VersionGUID.ToString();
            // 프로젝트 이름
            string filenameShort = Path.GetFileNameWithoutExtension(filename);
            // 프로젝트 GUID
            string creationGUID = doc.CreationGUID.ToString();

            // 24.5.14. 수정
            var startTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string JsonFile = Path.Combine(folderPath, DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + $"_{creationGUID}" + $"_{doc.Title}" + ".json");
            fileAndPath[$"{creationGUID}"] = JsonFile;
            JObject JStart = JObject.Parse(beginLog);
            fileAndJObject[$"{creationGUID}"] = JStart;
            JObject jobject = fileAndJObject[$"{creationGUID}"];

            File.WriteAllText(JsonFile, String.Empty);
            using (var streamWriter = new StreamWriter(JsonFile, true))
            {
                using (var writer = new JsonTextWriter(streamWriter))
                {
                    JObject bimLog = (JObject)jobject["bimlog"];
                    bimLog.Add("startTime", startTime);
                    bimLog.Property("startTime").AddAfterSelf(new JProperty("userName", userId));
                    bimLog.Property("userName").AddAfterSelf(new JProperty("projectName", filenameShort));
                    bimLog.Property("projectName").AddAfterSelf(new JProperty("projectGUID", creationGUID));
                    bimLog.Property("projectGUID").AddAfterSelf(new JProperty("Log", new JArray()));
                    var serializer = new JsonSerializer();
                    serializer.Formatting = Newtonsoft.Json.Formatting.Indented;
                    serializer.Serialize(writer, jobject);
                }
            }
        }
        // 24.5.14. 수정
        public void DocumentCreatedTracker( object sender, DocumentCreatedEventArgs e )
        {
            Document doc = e.Document;
            userId = doc.Application.Username;
            string filename = doc.PathName;
            Debug.Write(userId);
            string projectId = "";
            // 프로젝트 GUID

            // 24.5.16.
            string filenameShort = doc.Title;
            string creationGUID = doc.CreationGUID.ToString();

            var startTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string JsonFile = Path.Combine(folderPath, DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + $"_{creationGUID}" + $"_{doc.Title}" + ".json");
            fileAndPath[$"{creationGUID}"] = JsonFile;
            JObject JStart = JObject.Parse(beginLog);
            fileAndJObject[$"{creationGUID}"] = JStart;
            JObject jobject = fileAndJObject[$"{creationGUID}"];

            File.WriteAllText(JsonFile, String.Empty);
            using (var streamWriter = new StreamWriter(JsonFile, true))
            {
                using (var writer = new JsonTextWriter(streamWriter))
                {
                    JObject bimLog = (JObject)jobject["bimlog"];
                    bimLog.Add("startTime", startTime);
                    bimLog.Property("startTime").AddAfterSelf(new JProperty("userName", userId));
                    bimLog.Property("userName").AddAfterSelf(new JProperty("projectName", filenameShort));
                    bimLog.Property("projectName").AddAfterSelf(new JProperty("projectGUID", creationGUID));
                    bimLog.Property("projectGUID").AddAfterSelf(new JProperty("Log", new JArray()));
                    var serializer = new JsonSerializer();
                    serializer.Formatting = Newtonsoft.Json.Formatting.Indented;
                    serializer.Serialize(writer, jobject);
                    Debug.WriteLine("opened");
                }
            }
        }

        public void DocumentSavingTracker( object sender, DocumentSavingEventArgs e )
        {
            Document doc = e.Document;
            string filename = doc.PathName;
            // 프로젝트 GUID
            BasicFileInfo info = BasicFileInfo.Extract(filename);
            DocumentVersion v = info.GetDocumentVersion();
            string projectId = v.VersionGUID.ToString();

            string JsonFile = fileAndPath[$"{doc.CreationGUID}"];
            JObject jobject = fileAndJObject[$"{doc.CreationGUID}"];

            string index = folderPath + "\\" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + $"_{doc.CreationGUID}";
            string extension = JsonFile.Substring(0, index.Length);
            JsonFile = extension + $"_{doc.Title}_saved.json";
            var savedTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");


            File.WriteAllText(JsonFile, String.Empty);
            using (var streamWriter = new StreamWriter(JsonFile, true))
            {
                using (var writer = new JsonTextWriter(streamWriter))
                {
                    JObject bimLog = (JObject)jobject["bimlog"];
                    if (bimLog.ContainsKey("endTime"))
                    {
                        bimLog["endTime"] = savedTime;
                        if (!bimLog.ContainsKey("Saved"))
                        {
                            bimLog.Property("endTime").AddAfterSelf(new JProperty("Saved", true));
                        }
                    }

                    else
                    {
                        bimLog.Property("startTime").AddAfterSelf(new JProperty("endTime", savedTime));
                        bimLog.Property("endTime").AddAfterSelf(new JProperty("Saved", true));

                    }
                    var serializer = new JsonSerializer();
                    serializer.Formatting = Newtonsoft.Json.Formatting.Indented;
                    serializer.Serialize(writer, jobject);
                    //Debug.WriteLine("opened");
                }
            }
            SendData(jobject);

        }

        public void DocumentSavedAsTracker( object sender, DocumentSavedAsEventArgs e )
        {

            Document doc = e.Document;
            userId = doc.Application.Username;
            string filename = doc.PathName;
            // 프로젝트 GUID
            BasicFileInfo info = BasicFileInfo.Extract(filename);
            DocumentVersion v = info.GetDocumentVersion();
            string projectId = v.VersionGUID.ToString();
            string filenameShort = Path.GetFileNameWithoutExtension(filename);

            string JsonFile = fileAndPath[$"{doc.CreationGUID}"];
            string index = folderPath + "\\" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + $"_{doc.CreationGUID}";
            string extension = JsonFile.Substring(0, index.Length);
            JsonFile = extension + $"_{doc.Title}_saved.json";
            JObject jobject = fileAndJObject[$"{doc.CreationGUID}"];
            var savedTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            if (fileName != filenameShort)
            {
                File.WriteAllText(JsonFile, String.Empty);
                using (var streamWriter = new StreamWriter(JsonFile, true))
                {
                    using (var writer = new JsonTextWriter(streamWriter))
                    {
                        JObject bimLog = (JObject)jobject["bimlog"];
                        if (bimLog.ContainsKey("endTime"))
                        {
                            bimLog["endTime"] = savedTime;
                            bimLog.Property("endTime").AddAfterSelf(new JProperty("Saved", true));
                            bimLog["projectName"] = doc.Title;
                        }
                        else
                        {
                            bimLog["projectName"] = doc.Title;
                            bimLog.Property("startTime").AddAfterSelf(new JProperty("endTime", savedTime));
                            bimLog.Property("endTime").AddAfterSelf(new JProperty("Saved", true));

                        }
                        var serializer = new JsonSerializer();
                        serializer.Formatting = Newtonsoft.Json.Formatting.Indented;
                        serializer.Serialize(writer, jobject);
                    }
                }
            }
            SendData(jobject);

        }

        public void DocumentClosingTracker( object sender, DocumentClosingEventArgs e )
        {
            Document doc = e.Document;
            userId = doc.Application.Username;
            // 프로젝트 GUID
            string JsonFile = fileAndPath[$"{doc.CreationGUID}"];
            JObject jobject = fileAndJObject[$"{doc.CreationGUID}"];

            var endTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            // JSON
            // 24.5.17. 수정
            var saved = jobject["bimlog"]?["Saved"]?.ToString();
            string fileName = Path.GetFileNameWithoutExtension(JsonFile);
            Debug.WriteLine(fileName);
            string[] parts = fileName.Split('_');
            string realName = "";
            for (int i = 3 ; i < parts.Length ; i++)
            {
                realName += parts[i];
                realName += "_";
            }
            Debug.WriteLine(realName);
            realName = realName.TrimEnd('_');
            Debug.WriteLine(realName);
            if (saved == null || saved == "False")
            {
                File.WriteAllText(JsonFile, String.Empty);
                using (var streamWriter = new StreamWriter(JsonFile, true))
                {
                    using (var writer = new JsonTextWriter(streamWriter))
                    {
                        JObject bimLog = (JObject)jobject["bimlog"];
                        bimLog["projectName"] = realName;
                        if (bimLog.ContainsKey("endTime"))
                        {
                            bimLog["endTime"] = endTime;
                        }
                        else
                        {
                            bimLog.Property("startTime").AddAfterSelf(new JProperty("endTime", endTime));
                            bimLog.Property("endTime").AddAfterSelf(new JProperty("Saved", false));

                        }
                        var serializer = new JsonSerializer();
                        serializer.Formatting = Newtonsoft.Json.Formatting.Indented;
                        serializer.Serialize(writer, jobject);
                    }
                }
                SendData(jobject);
            }
            else
            {
                try
                {
                    if (File.Exists(JsonFile))
                    {
                        File.Delete(JsonFile);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("File Delete err: " + ex);
                }
            }
        }

        private void CommandExecuted( object sender, RibbonItemExecutedEventArgs args )
        {
            try
            {
                Autodesk.Windows.RibbonItem it = args.Item;
                if (args != null)
                {
                    // 수정
                    //CommandLog commandLog = new CommandLog() { Timestamp = DateTime.Now.ToString(), ElementId = userId, CommandType = "COMMAND", ElementCategory = it.ToString(), CommandId = it.Id, CommandCookie = it.Cookie };
                    //**JSON 작성 필요
                    //
                    //
                    //
                }
            }

            catch (Exception ex)
            {
                Autodesk.Revit.UI.TaskDialog.Show("ADD-IN FAILED", ex.Message);
            }
        }

        public void DocumentChangeTracker( object sender, DocumentChangedEventArgs args )
        {
            var app = sender as Autodesk.Revit.ApplicationServices.Application;   //send를 받아서 app에 어플리케이션으로 저장
            UIApplication uiapp = new UIApplication(app); // app에 대해 uiapplication으로 생성
            UIDocument uidoc = uiapp.ActiveUIDocument; //uiapp의 활성화된 UI다큐먼트
            Document doc = uidoc.Document; // uidoc의 다큐먼트
            //24.5.16. 수정
            string filename = doc.PathName;
            JArray jarray = new JArray();

            string JsonFile = fileAndPath[$"{doc.CreationGUID}"];
            JObject jobject = fileAndJObject[$"{doc.CreationGUID}"];
            var checkSaved = jobject["bimlog"]?["Saved"]?.ToString();
            if (checkSaved != null && checkSaved == "True")
            {
                jobject["bimlog"]["Saved"] = false;
            }
            // 프로젝트 이름
            string filenameShort = Path.GetFileNameWithoutExtension(filename);

            // 이정헌 24.5.17.
            UndoOperation operationType = args.Operation;
            #region undoOperation
            if (operationType == UndoOperation.TransactionUndone)
            {
                var timestamp = DateTime.Now.ToString();
                JObject addJarray = new JObject(new JProperty("Undone", new JObject()));
                JObject undoTime = (JObject)addJarray["Undone"];
                undoTime.Add("UndoneTime", timestamp);

                if (addJarray != null)
                {
                    File.WriteAllText(JsonFile, String.Empty);
                    using (var streamWriter = new StreamWriter(JsonFile, true))
                    {
                        using (var writer = new JsonTextWriter(streamWriter))
                        {
                            JObject o = (JObject)JToken.FromObject(addJarray);
                            jarray.Add(o);
                            JArray JLog = (JArray)jobject["bimlog"]["Log"];
                            JLog.Merge(jarray, new JsonMergeSettings { MergeArrayHandling = MergeArrayHandling.Union });
                            var json = new JObject();
                            var serializer = new JsonSerializer();
                            serializer.Formatting = Formatting.Indented;
                            serializer.Serialize(writer, jobject);
                        }
                    }
                }
                return;
            }
            else if (operationType == UndoOperation.TransactionRedone)
            {
                var timestamp = DateTime.Now.ToString();
                JObject addJarray = new JObject(new JProperty("Redone", new JObject()));
                JObject undoTime = (JObject)addJarray["Redone"];
                undoTime.Add("RedoneTime", timestamp);
                JArray jLog = (JArray)jobject["bimlog"]["Log"];
                JToken test = jLog.Last;
                Debug.WriteLine(test);
                if (addJarray != null)
                {
                    File.WriteAllText(JsonFile, String.Empty);
                    using (var streamWriter = new StreamWriter(JsonFile, true))
                    {
                        using (var writer = new JsonTextWriter(streamWriter))
                        {
                            JObject o = (JObject)JToken.FromObject(addJarray);
                            jarray.Add(o);
                            JArray JLog = (JArray)jobject["bimlog"]["Log"];
                            JLog.Merge(jarray, new JsonMergeSettings { MergeArrayHandling = MergeArrayHandling.Union });
                            var json = new JObject();
                            var serializer = new JsonSerializer();
                            serializer.Formatting = Formatting.Indented;
                            serializer.Serialize(writer, jobject);
                        }
                    }
                }
                return;
            }
            #endregion

            Selection sel = uidoc.Selection;
            ICollection<ElementId> selectedIds = sel.GetElementIds();
            ICollection<ElementId> deletedElements = args.GetDeletedElementIds();
            ICollection<ElementId> modifiedElements = args.GetModifiedElementIds();
            ICollection<ElementId> addedElements = args.GetAddedElementIds();
            int counter = deletedElements.Count + modifiedElements.Count + addedElements.Count;

            if (modifiedElements.Count != 0)
            {
                // 24.6.24. 작업 체크
                string cmd = "MODIFIED";
                foreach (ElementId eid in modifiedElements)
                {
                    var elem = doc.GetElement(eid);
                    Debug.WriteLine("Modified Element:  " + elem);
                    try
                    {
                        if (selectedIds.Contains(eid) || selectedElementList.Contains(eid) || addedElementList.Contains(eid) || addedStairList.Contains(eid))
                        // 선택된 객체랑 다른 객체가 수정되는 경우가 존재.
                        // 예를들어 profile로 형상을 표현하는 객체들.
                        {
                            if (doc.GetElement(eid) == null || doc.GetElement(doc.GetElement(eid).GetTypeId()) == null)
                                continue;
                            dynamic logdata = Log(doc, cmd, eid);
                            //JSON
                            // 24.3.18. 수정
                            if (logdata != null)
                            {
                                File.WriteAllText(JsonFile, String.Empty);
                                using (var streamWriter = new StreamWriter(JsonFile, true))
                                {
                                    using (var writer = new JsonTextWriter(streamWriter))
                                    {
                                        JObject o = (JObject)JToken.FromObject(logdata);
                                        jarray.Add(o);
                                        JArray JLog = (JArray)jobject["bimlog"]["Log"];
                                        JLog.Merge(jarray, new JsonMergeSettings { MergeArrayHandling = MergeArrayHandling.Union });
                                        var json = new JObject();
                                        var serializer = new JsonSerializer();
                                        serializer.Formatting = Newtonsoft.Json.Formatting.Indented;
                                        serializer.Serialize(writer, jobject);
                                    }
                                }
                            }
                        }

                    }
                    catch { }
                }
            }
            if (addedElements.Count != 0)
            {
                string cmd = "ADDED";

                foreach (ElementId eid in addedElements)
                {
                    var elem = doc.GetElement(eid);
                    //Debug.Print(doc.GetElement(eid).Category.ToString());
                    Debug.WriteLine("Added Element:  " + elem);

                    try
                    {
                        if (doc.GetElement(eid) == null || doc.GetElement(doc.GetElement(eid).GetTypeId()) == null)
                            continue;
                        // 24.7.1. 수정
                        // AMM 객체 
                        if (!CheckAddedElement(eid, sender))
                            continue;
                        dynamic logdata = Log(doc, cmd, eid);
                        //JSON
                        // 24.3.18. 수정
                        if (logdata != null)
                        {
                            File.WriteAllText(JsonFile, String.Empty);
                            using (var streamWriter = new StreamWriter(JsonFile, true))
                            {
                                using (var writer = new JsonTextWriter(streamWriter))
                                {
                                    JObject o = (JObject)JToken.FromObject(logdata);
                                    jarray.Add(o);
                                    JArray JLog = (JArray)jobject["bimlog"]["Log"];
                                    JLog.Merge(jarray, new JsonMergeSettings
                                    {
                                        MergeArrayHandling = MergeArrayHandling.Union
                                    });
                                    var json = new JObject();
                                    var serializer = new JsonSerializer();
                                    serializer.Formatting = Newtonsoft.Json.Formatting.Indented;
                                    serializer.Serialize(writer, jobject);

                                }
                            }
                        }
                    }
                    catch { }
                }
            }
            if (deletedElements.Count != 0)
            {

                string cmd = "DELETED";

                foreach (ElementId eid in deletedElements)
                {
                    try
                    {

                        dynamic logdata = Log(doc, cmd, eid);
                        if (logdata != null)
                        {
                            File.WriteAllText(JsonFile, String.Empty);
                            using (var streamWriter = new StreamWriter(JsonFile, true))
                            {
                                using (var writer = new JsonTextWriter(streamWriter))
                                {
                                    JObject data = (JObject)JToken.FromObject(logdata);
                                    jarray.Add(data);
                                    JArray JLog = (JArray)jobject["bimlog"]["Log"];
                                    JLog.Merge(jarray, new JsonMergeSettings { MergeArrayHandling = MergeArrayHandling.Union });
                                    var json = new JObject();
                                    var serializer = new JsonSerializer();
                                    serializer.Formatting = Newtonsoft.Json.Formatting.Indented;
                                    serializer.Serialize(writer, jobject);
                                }
                            }
                        }

                    }
                    catch { }
                }
            }
        }

        public void SelectionChangeTracker( object sender, SelectionChangedEventArgs e )
        {
            Document doc = e.GetDocument();
            var selected = e.GetSelectedElements();
            foreach (ElementId eId in selected)
            {
                var elem = doc.GetElement(eId);
                var cat = elem.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString();
                var checkSketch = elem.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsElementId();
                // 이거임!!
                Category checkCategory = Category.GetCategory(doc, checkSketch);
                ㅈDebug.WriteLine(checkCategory.Name);
                // 여기에 null이면 안됨 유의!!!
                //  ex wall opening
                if (checkCategory.Name == "<Sketch>")
                {
                    Debug.WriteLine("SketchElement");
                    continue;
                }
                else
                {
                    selectedElementList.Clear();
                    selectedElementList.Add(eId);
                }
            }
        }

        private void FailureTracker( object sender, FailuresProcessingEventArgs e )
        {
            var app = sender as Autodesk.Revit.ApplicationServices.Application;
            UIApplication uiapp = new UIApplication(app);

            UIDocument uidoc = uiapp.ActiveUIDocument;
            if (uidoc != null)
            {
                Document doc = uidoc.Document;
                string user = doc.Application.Username;
                string filename = doc.PathName;
                string filenameShort = Path.GetFileNameWithoutExtension(filename);

                FailuresAccessor failuresAccessor = e.GetFailuresAccessor();
                IList<FailureMessageAccessor> fmas = failuresAccessor.GetFailureMessages();


                if (fmas.Count != 0)
                {
                    foreach (FailureMessageAccessor fma in fmas)
                    {
                        // 수정
                        //FailureLog failureLog = new FailureLog() { Timestamp = DateTime.Now.ToString(), ElementId = user, CommandType = "FAILURE", ElementCategory = failuresAccessor.GetTransactionName(), FailureMessage = failuresAccessor.GetFailureMessages().ToString() };
                        //**JSON 작성 필요
                        //
                        //
                        //
                    }
                }

            }

        }
        #endregion
        #region Functions
        public void GetParameter( Parameter p, JObject addJson )
        {
            Parameter param = p;
            JObject builtin = (JObject)addJson["Parameter"]["Built-In"];
            JObject custom = (JObject)addJson["Parameter"]["Custom"];
            var pName = param.Definition.Name;
            var pDef = param.Definition as InternalDefinition;
            var pDefName = pDef.BuiltInParameter;
            bool checkNull = p.HasValue;
            var pAsValueString = p.AsValueString();
            var storageType = param.StorageType;
            string pStorageType = "";
            if (checkNull == true)
            {
                if (storageType == StorageType.String)
                {
                    pStorageType = "String";
                    var pValue = param.AsString();
                    Debug.WriteLine($"{pName} as definition: {pDefName} : {pAsValueString}, storageType: {pStorageType}, Value: {pValue}");
                    if (pDefName != BuiltInParameter.INVALID)
                    {
                        // Built-in Parameter
                        JObject parameters = new JObject(
                            new JProperty("StorageType", pStorageType),
                            new JProperty("Value", pValue),
                            new JProperty("ValueString", pAsValueString)
                        );
                        builtin.Add($"{pDefName}", parameters);
                    }
                    else
                    {
                        // Custom parameter
                        JObject parameters = new JObject(
                                                    new JProperty("StorageType", pStorageType),
                                                    new JProperty("Value", pValue),
                                                    new JProperty("ValueString", pAsValueString)
                                                );
                        custom.Add($"{pName}", parameters);
                    }
                }
                else if (storageType == StorageType.Double)
                {
                    pStorageType = "Double";
                    var pValue = param.AsDouble();
                    Debug.WriteLine($"{pName} as definition: {pDefName} : {pAsValueString}, storageType: {pStorageType}, Value: {pValue}");
                    if (pDefName != BuiltInParameter.INVALID)
                    {
                        // Built-in Parameter
                        JObject parameters = new JObject(
                            new JProperty("StorageType", pStorageType),
                            new JProperty("Value", pValue),
                            new JProperty("ValueString", pAsValueString)
                        );
                        builtin.Add($"{pDefName}", parameters);
                    }
                    else
                    {
                        // Custom parameter
                        JObject parameters = new JObject(
        new JProperty("StorageType", pStorageType),
                            new JProperty("Value", pValue),
                            new JProperty("ValueString", pAsValueString)
                        );
                        custom.Add($"{pName}", parameters);
                    }
                }
                else if (storageType == StorageType.Integer)
                {
                    pStorageType = "Integer";
                    var pValue = param.AsInteger();
                    Debug.WriteLine($"{pName} as definition: {pDefName} : {pAsValueString}, storageType: {pStorageType}, Value: {pValue}");
                    if (pDefName != BuiltInParameter.INVALID)
                    {
                        // Built-in Parameter
                        JObject parameters = new JObject(
                            new JProperty("StorageType", pStorageType),
                            new JProperty("Value", pValue),
                            new JProperty("ValueString", pAsValueString)
                        );
                        builtin.Add($"{pDefName}", parameters);
                    }
                    else
                    {
                        // Custom parameter
                        JObject parameters = new JObject(
        new JProperty("StorageType", pStorageType),
        new JProperty("Value", pValue),
                            new JProperty("ValueString", pAsValueString)
                        );
                        custom.Add($"{pName}", parameters);
                    }
                }
                else if (storageType == StorageType.ElementId)
                {
                    pStorageType = "ElementId";
                    var pValue = param.AsElementId().Value;
                    Debug.WriteLine($"{pName} as definition: {pDefName} : {pAsValueString}, storageType: {pStorageType}, Value: {pValue}");
                    if (pDefName != BuiltInParameter.INVALID)
                    {
                        // Built-in Parameter
                        JObject parameters = new JObject(
                            new JProperty("StorageType", pStorageType),
                            new JProperty("Value", pValue),
                            new JProperty("ValueString", pAsValueString)
                        );
                        builtin.Add($"{pDefName}", parameters);
                    }
                    else
                    {
                        // Custom parameter
                        JObject parameters = new JObject(
                            new JProperty("StorageType", pStorageType),
        new JProperty("Value", pValue),
                            new JProperty("ValueString", pAsValueString)
                        );
                        custom.Add($"{pName}", parameters);
                    }
                }
                else if (storageType == StorageType.None)
                {
                    pStorageType = "None";
                    var pValue = param.AsString();
                    Debug.WriteLine($"{pName} as definition: {pDefName} : {pAsValueString}, storageType: {pStorageType}, Value: {pValue}");
                    if (pDefName != BuiltInParameter.INVALID)
                    {
                        // Built-in Parameter
                        JObject parameters = new JObject(
                            new JProperty("StorageType", pStorageType),
                            new JProperty("Value", pValue),
                            new JProperty("ValueString", pAsValueString)
                        );
                        builtin.Add($"{pDefName}", parameters);
                    }
                    else
                    {
                        // Custom parameter
                        JObject parameters = new JObject(
                            new JProperty("StorageType", pStorageType),
                            new JProperty("Value", pValue),
        new JProperty("ValueString", pAsValueString)
                        );
                        custom.Add($"{pName}", parameters);
                    }
                }

            }


        }
        // 24.7.17. 수정
        public void GetLayer( Element elem, JObject addJObject, Document doc )
        {
            JArray LayerJObject = (JArray)addJObject["Layers"];
            var test = doc.GetElement(elem.GetTypeId()) as HostObjAttributes;
            var testlist = test.GetCompoundStructure().GetLayers();
            if (testlist.Count != 0)
            {

                foreach (CompoundStructureLayer layer in testlist)
                {
                    var func = layer.Function.ToString();
                    var width = layer.Width;
                    var material_id = layer.MaterialId.Value;
                    string material_name = "None";
                    if (material_id != -1)
                    {
                        material_name = doc.GetElement(layer.MaterialId).Name;
                    }
                    JObject layerInfo = new JObject(
                        new JProperty("Function", func),
                        new JProperty("Width", width),
                        new JProperty("MaterialId", material_id),
                        new JProperty("MaterialName", material_name));
                    LayerJObject.Add(layerInfo);
                }
            }
            else
            {
                addJObject.Property("Layers").Remove();
            }

        }
        public void FamilyGetLayer( Element elem, JObject addJObject, Document doc )
        {
            JArray LayerJObject = (JArray)addJObject["Layers"];
            var test = doc.GetElement(elem.GetTypeId()).GetMaterialIds(elem.Category.HasMaterialQuantities);
            if (test.Count != 0)
            {
                foreach (ElementId materialId in test)
                {
                    var material_name = doc.GetElement(materialId).Name.ToString();
                    JObject layerInfo = new JObject(
                    new JProperty("MaterialId", materialId),
                    new JProperty("MaterialName", material_name));
                    LayerJObject.Add(layerInfo);
                }
            }
            else
            {
                addJObject.Property("Layers").Remove();
            }
        }
        public bool CheckAddedElement( ElementId eid, object sender )
        {
            bool checkAddedElement = false;
            var app = sender as Autodesk.Revit.ApplicationServices.Application;   //send를 받아서 app에 어플리케이션으로 저장
            UIApplication uiapp = new UIApplication(app); // app에 대해 uiapplication으로 생성
            UIDocument uidoc = uiapp.ActiveUIDocument; //uiapp의 활성화된 UI다큐먼트
            Document doc = uidoc.Document; // uidoc의 다큐먼트

            var element = doc.GetElement(eid);
            if (element.Category != null)
            {
                // 24.7.3. Stairs 의 경우 분리해서 생각해야됨.
                // 24.7.15. 수정
                //var cat = element.Category.Name;
                var cat = element.Category.BuiltInCategory.ToString();
                if (cat == "OST_Roofs" || cat == "OST_StairsRailing" || cat == "OST_Floors" || cat == "OST_Ceilings")
                {
                    checkAddedElement = true;
                    addedElementList.Add(eid);
                }
                else if (cat == "OST_Stairs" || cat == "OST_StairsRuns" || cat == "OST_StairsLandings")
                {
                    checkAddedElement = true;
                    addedStairList.Add(eid);
                }
                else
                {
                    checkAddedElement = true;
                }
            }
            return checkAddedElement;

        }
        #endregion
        #region Elements Log
        public dynamic Log( Document doc, String cmd, ElementId eid )
        {
            JObject addJarray = new JObject(
                new JProperty("Common", new JObject()),
                new JProperty("Geometry", new JObject()),
                new JProperty("Parameter", new JObject(
                    new JProperty("Built-In", new JObject()),
                    new JProperty("Custom", new JObject()))),
                new JProperty("Property", new JObject()),
                // 24.7.17. 수정
                new JProperty("Layers", new JArray())
                );
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            var id = eid.ToString();
            var elem = doc.GetElement(eid);

            if (cmd == "DELETED")
            {
                //var log = new GeneralLog();
                //log.Timestamp = timestamp;
                //log.ElementId = id;
                //log.CommandType = cmd;
                //Common
                JObject deleteCommon = (JObject)addJarray["Common"];
                deleteCommon.Add("CommandType", cmd);
                deleteCommon.Add("Timestamp", timestamp);
                deleteCommon.Add("ElementId", id);
                addJarray.Property("Geometry").Remove();
                addJarray.Property("Parameter").Remove();
                addJarray.Property("Property").Remove();
                return addJarray;
            }
            else
            {
                if (elem.Category != null)
                {
                    var cat = elem.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString();
                    var fam = elem.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();
                    var typ = elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString();

                    // 24.7.15. 수정
                    var catForSwitch = elem.Category.BuiltInCategory.ToString();
                    Debug.WriteLine(catForSwitch);
                    //var roomIds = GetRooms(elem);
                    var worksetId = elem.WorksetId.ToString();

                    JObject common = (JObject)addJarray["Common"];
                    common.Add("Timestamp", timestamp);
                    common.Add("ElementId", id);
                    common.Add("CommandType", cmd);
                    common.Add("ElementCategory", cat);
                    common.Add("ElementFamily", fam);
                    common.Add("ElementType", typ);
                    switch (catForSwitch)
                    {
                        case "OST_Walls":
                            GetLayer(elem, addJarray, doc);
                            var wall = elem as Wall;
                            var norm = wall.Orientation.ToString();
                            var isStr = wall.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT).AsInteger();
                            bool wallIsProfileWall = true;
                            var wallCurve = new JObject();
                            List<Autodesk.Revit.DB.Curve> curves = new List<Autodesk.Revit.DB.Curve>();

                            //24.5.1. 
                            // Parameter 부분
                            var wallParameters = wall.Parameters;
                            foreach (Parameter p in wallParameters)
                            {
                                GetParameter(p, addJarray);
                            }

                            // geometry
                            if (doc.GetElement(wall.SketchId) == null)
                            {
                                wallIsProfileWall = false;
                            }
                            if (wallIsProfileWall == true)
                            {
                                Sketch wallSketch = doc.GetElement(wall.SketchId) as Sketch;
                                //wallCurve = GetProfileDescription(wallSketch);
                            }
                            else
                            {
                                Curve wallLocCrv = (wall.Location as LocationCurve).Curve;
                                wallCurve = GetCurveDescription(wallLocCrv);
                            }

                            var wallFlipped = wall.Flipped;
                            // 24.7.15. 수정
                            //var wallWidth = doc.GetElement(wall.GetTypeId()).GetParameters("Width").First().AsValueString();
                            var wallWidth = doc.GetElement(wall.GetTypeId()).get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM).AsValueString();
                            //24.3.19. 수정
                            // geometry
                            JObject wallGeometry = (JObject)addJarray["Geometry"];
                            wallGeometry.Add("IsProfileWall", wallIsProfileWall);
                            wallGeometry.Add("Curve", wallCurve);
                            // property
                            JObject wallProperty = (JObject)addJarray["Property"];
                            wallProperty.Add("Flipped", wallFlipped);
                            wallProperty.Add("Width", wallWidth);

                            return addJarray;

                        case "OST_Floors":
                            // 24.7.2. 수정
                            GetLayer(elem, addJarray, doc);

                            var floor = elem as Floor;
                            var floorCheckAdded = floor.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsValueString();
                            if (floorCheckAdded == null)
                            {
                                if (addedElementList.Contains(eid))
                                {
                                    return null;
                                }
                            }
                            else
                            {
                                if (addedElementList.Contains(eid))
                                {
                                    common["CommandType"] = "ADDED";
                                    addedElementList.Remove(eid);
                                }
                                JObject floorGeometry = (JObject)addJarray["Geometry"];
                                JObject floorProperty = (JObject)addJarray["Property"];

                                var floorSketch = (doc.GetElement(floor.SketchId) as Sketch);
                                var profileEmpty = floorSketch.Profile.IsEmpty;
                                var floorEIds = floorSketch.GetAllElements();
                                JObject floorSlopeArrow = new JObject();
                                string floorSlope = null;
                                JObject floorSpanDirection = new JObject();
                                foreach (ElementId floorEId in floorEIds)
                                {
                                    // 24.7.15. 수정
                                    // 경사 화살표 
                                    if (doc.GetElement(floorEId).Name == "Slope Arrow" || doc.GetElement(floorEId).Name == "경사 화살표")
                                    {
                                        var Floorcrv = (doc.GetElement(floorEId) as CurveElement).GeometryCurve;
                                        floorSlopeArrow = GetCurveDescription(Floorcrv);

                                        List<Parameter> floorParams = (doc.GetElement(floorEId) as ModelLine).GetOrderedParameters().ToList();

                                        foreach (var param in floorParams)
                                        {
                                            var defslop = param.Definition.ToString();
                                            Debug.WriteLine(defslop);
                                            if (param.Definition.Name == "Slope" || param.Definition.Name == "경사")
                                            // 24.7.15. 수정
                                            // 경사 ROOF_SLOPE
                                            {
                                                floorSlope = param.AsValueString();
                                            }
                                        }
                                    }
                                    if (doc.GetElement(floorEId).Name == "Span Direction Edges" || doc.GetElement(floorEId).Name == "스팬 방향 모서리")
                                    {
                                        var Floorcrv = (doc.GetElement(floorEId) as CurveElement).GeometryCurve;
                                        floorSpanDirection = GetCurveDescription(Floorcrv);
                                    }
                                }
                                var floorProfile = GetProfileDescription(floorSketch);
                                floorGeometry.Add("Profile", floorProfile);
                                if (floorSlopeArrow.HasValues == false)
                                {
                                    floorGeometry.Add("SlopeArrow", "None");
                                }
                                else
                                {
                                    floorGeometry.Add("SlopeArrow", floorSlopeArrow);
                                    floorGeometry.Add("SlopeAngle", floorSlope);
                                }
                                if (floorSpanDirection.HasValues == false)
                                {
                                    floorGeometry.Add("SpanDirection", "None");
                                }
                                else
                                {
                                    floorGeometry.Add("SpanDirection", floorSpanDirection);
                                }

                                // floor Parameter
                                var floorParameters = floor.Parameters;
                                foreach (Parameter p in floorParameters)
                                {
                                    GetParameter(p, addJarray);
                                }
                            }
                            return addJarray;

                        case "OST_Roofs":

                            if (elem.GetType().Name == "FootPrintRoof")
                            {
                                cat = "Roofs: FootPrintRoof";
                                GetLayer(elem, addJarray, doc);
                                var footprintroof = elem as FootPrintRoof;
                                var fpRoofCheckAdded = footprintroof.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsValueString();
                                Debug.WriteLine(fpRoofCheckAdded);
                                if (fpRoofCheckAdded == null)
                                {
                                    if (addedElementList.Contains(eid))
                                    {
                                        return null;
                                    }
                                }
                                else
                                {
                                    if (addedElementList.Contains(eid))
                                    {
                                        common["CommandType"] = "ADDED";
                                        addedElementList.Remove(eid);
                                    }
                                    common["ElementCategory"] = cat;
                                    JObject fpRoofGeometry = (JObject)addJarray["Geometry"];
                                    ElementClassFilter roofSktFilter = new ElementClassFilter(typeof(Sketch));
                                    Sketch footPrintRoofSketch = doc.GetElement(footprintroof.GetDependentElements(roofSktFilter).ToList()[0]) as Sketch;
                                    var footprintroofFootPrint = GetProfileDescription(footPrintRoofSketch);
                                    fpRoofGeometry.Add("FootPrint", footprintroofFootPrint);

                                    var footprintroofParameters = footprintroof.Parameters;
                                    foreach (Parameter p in footprintroofParameters)
                                    {
                                        GetParameter(p, addJarray);
                                    }
                                }
                            }
                            else if (elem.GetType().Name == "ExtrusionRoof")
                            {
                                cat = "Roofs: ExtrusionRoof";
                                GetLayer(elem, addJarray, doc);

                                var extrusionroof = elem as ExtrusionRoof;
                                var eRoofCheckAdded = extrusionroof.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsValueString();
                                Debug.WriteLine(eRoofCheckAdded);
                                if (eRoofCheckAdded == null)
                                {
                                    if (addedElementList.Contains(eid))
                                    {
                                        return null;
                                    }
                                }
                                else
                                {
                                    if (addedElementList.Contains(eid))
                                    {
                                        common["CommandType"] = "ADDED";
                                        addedElementList.Remove(eid);
                                    }
                                    common["ElementCategory"] = cat;
                                    JObject eRoofGeometry = (JObject)addJarray["Geometry"];
                                    var extrusionroofCrvLoop = new CurveLoop();
                                    var extrusionroofProfileCurves = extrusionroof.GetProfile();
                                    foreach (ModelCurve curve in extrusionroofProfileCurves)
                                    {
                                        extrusionroofCrvLoop.Append(curve.GeometryCurve);
                                    }
                                    var extrusionroofWorkPlane = GetPlaneDescription(extrusionroofCrvLoop.GetPlane());
                                    var extrusionroofProfile = GetCurveLoopDescription(extrusionroofCrvLoop);

                                    eRoofGeometry.Add("WorkPlane", extrusionroofWorkPlane);
                                    eRoofGeometry.Add("Profile", extrusionroofProfile);

                                    var extrusionroofParameter = extrusionroof.Parameters;
                                    foreach (Parameter p in extrusionroofParameter)
                                    {
                                        GetParameter(p, addJarray);
                                    }
                                }
                            }
                            else
                            {
                                return null;
                            }

                            return addJarray;

                        case "OST_Ceilings":
                            // 24.4.2. 수정
                            GetLayer(elem, addJarray, doc);
                            var ceiling = elem as Ceiling;

                            var ceilingCheckAdded = ceiling.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsValueString();
                            if (ceilingCheckAdded == null)
                            {
                                if (addedElementList.Contains(eid))
                                {
                                    return null;
                                }
                            }
                            else
                            {
                                if (addedElementList.Contains(eid))
                                {
                                    common["CommandType"] = "ADDED";
                                    addedElementList.Remove(eid);
                                }
                                var ceilingCurveLoops = GetProfileDescription((doc.GetElement(ceiling.SketchId) as Sketch));
                                var ceilingEIds = (doc.GetElement(ceiling.SketchId) as Sketch).GetAllElements();
                                JObject ceilingSlopeArrow = new JObject();
                                string ceilingSlope = null;

                                foreach (ElementId ceilingEId in ceilingEIds)
                                {
                                    if (doc.GetElement(ceilingEId).Name == "Slope Arrow" || doc.GetElement(ceilingEId).Name == "경사 화살표")
                                    {
                                        var ceilingCrv = (doc.GetElement(ceilingEId) as CurveElement).GeometryCurve;
                                        List<Parameter> slopeParams = (doc.GetElement(ceilingEId) as ModelLine).GetOrderedParameters().ToList();

                                        foreach (var param in slopeParams)
                                        {
                                            if (param.Definition.Name == "Slope" || param.Definition.Name == "경사")
                                            {
                                                ceilingSlope = param.AsValueString();
                                            }
                                        }

                                        ceilingSlopeArrow = GetCurveDescription(ceilingCrv);
                                    }
                                }
                                // parameter
                                var ceilingParameters = ceiling.Parameters;
                                foreach (Parameter p in ceilingParameters)
                                {
                                    GetParameter(p, addJarray);
                                }
                                var ceilingThickness = doc.GetElement(ceiling.GetTypeId()).get_Parameter(BuiltInParameter.CEILING_THICKNESS).AsDouble();
                                // geometry
                                JObject ceilingGeometry = (JObject)addJarray["Geometry"];
                                if (ceilingSlopeArrow.HasValues == false)
                                {
                                    ceilingGeometry.Add("SlopeArrow", "None");
                                }
                                else
                                {
                                    ceilingGeometry.Add("SlopeArrow", ceilingSlopeArrow);
                                    ceilingGeometry.Add("SlopeAngle", ceilingSlope);
                                }
                                ceilingGeometry.Add("CurveLoops", ceilingCurveLoops);
                                // property
                                JObject ceilingProperty = (JObject)addJarray["Property"];
                                ceilingProperty.Add("Thickness", ceilingThickness);

                            }


                            return addJarray;

                        case "OST_Levels":
                            // 24.3.29. 수정 , 문제
                            var level = elem as Level;

                            // property
                            // parameter
                            var levelParameters = level.Parameters;
                            foreach (Parameter p in levelParameters)
                            {
                                GetParameter(p, addJarray);
                            }

                            var levelElevation = level.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble();
                            // geometry
                            JObject levelGeometry = (JObject)addJarray["Geometry"];
                            levelGeometry.Add("Elevation", levelElevation);
                            return addJarray;

                        case "OST_Grids":

                            var grid = elem as Grid;
                            var gridCurve = GetCurveDescription(grid.Curve);
                            JObject gridGeometry = (JObject)addJarray["Geometry"];
                            gridGeometry.Add("Curve", gridCurve);
                            var gridParameters = grid.Parameters;
                            foreach (Parameter p in gridParameters)
                            {
                                GetParameter(p, addJarray);
                            }

                            return addJarray;

                        case "OST_Stairs":

                            var stair = elem as Stairs;
                            var checkAddedStair = stair.IsInEditMode();
                            if (checkAddedStair)
                            {
                                if (addedStairList.Contains(eid))
                                {
                                    return null;
                                }
                            }
                            else
                            {
                                if (addedStairList.Contains(eid))
                                {
                                    common["CommandType"] = "ADDED";
                                    addedStairList.Remove(eid);
                                }
                                var stairParameters = stair.Parameters;
                                foreach (Parameter p in stairParameters)
                                {
                                    GetParameter(p, addJarray);
                                }
                            }


                            return addJarray;

                        case "OST_StairsRuns":

                            var stairsruns = elem as StairsRun;
                            var runsHost = stairsruns.GetStairs();
                            var checkAddedRuns = runsHost.IsInEditMode();
                            if (checkAddedRuns)
                            {
                                return null;
                            }
                            else
                            {
                                if (addedStairList.Contains(eid))
                                {
                                    common["CommandType"] = "ADDED";
                                    addedStairList.Remove(eid);
                                }
                                // 여기에 코드 작성
                                var stairrunsStairId = stairsruns.GetStairs().Id.Value;
                                var stairsrunsLocationPath = GetCurveDescription(stairsruns.GetStairsPath().First());
                                JObject stairsGeometry = (JObject)addJarray["Geometry"];
                                stairsGeometry.Add("StairsId", stairrunsStairId);
                                stairsGeometry.Add("LocationPath", stairsrunsLocationPath);
                                var stairrunsParameters = stairsruns.Parameters;
                                foreach (Parameter p in stairrunsParameters)
                                {
                                    GetParameter(p, addJarray);
                                }
                            }

                            return addJarray;

                        case "OST_StairsLandings":

                            var stairslandings = elem as StairsLanding;
                            var landingsHost = stairslandings.GetStairs();
                            var checkAddedLandings = landingsHost.IsInEditMode();
                            if (checkAddedLandings)
                            {
                                return null;
                            }
                            else
                            {
                                if (addedStairList.Contains(eid))
                                {
                                    common["CommandType"] = "ADDED";
                                    addedStairList.Remove(eid);
                                }
                                // 여기 코드 작성
                                var stairslandingsStairsId = stairslandings.GetStairs().Id.ToString();
                                var stairslandingsCurveLoop = GetCurveLoopDescription(stairslandings.GetFootprintBoundary());
                                JObject stairslandingsGeometry = (JObject)addJarray["Geometry"];
                                stairslandingsGeometry.Add("StairsId", stairslandingsStairsId);
                                stairslandingsGeometry.Add("CuverLoop", stairslandingsCurveLoop);
                                var stairslandingsParameter = stairslandings.Parameters;
                                foreach (Parameter p in stairslandingsParameter)
                                {
                                    GetParameter(p, addJarray);
                                }
                            }

                            return addJarray;

                        case "OST_StairsRailing":

                            var railing = elem as Railing;
                            var checkHost = railing.HasHost;
                            var railingsCheckAdded = railing.TopRail.Value;
                            if (railingsCheckAdded == -1)
                            {
                                if (addedElementList.Contains(eid))
                                {
                                    return null;
                                }
                            }
                            else
                            {
                                if (checkHost)
                                {
                                    var rCat = railing.Category.Name;
                                    if (rCat == "난간")
                                    {
                                        common["ElementCategory"] = "계단: 계단난간";
                                    }
                                    else
                                    {
                                        common["ElementCategory"] = "Stairs: Railings";
                                        if (addedElementList.Contains(eid))
                                        {
                                            addedElementList.Remove(eid);
                                        }
                                    }
                                }
                                if (addedElementList.Contains(eid))
                                {
                                    common["CommandType"] = "ADDED";
                                    addedElementList.Remove(eid);
                                }
                                // Geometry
                                JObject railingGeometry = (JObject)addJarray["Geometry"];
                                ElementClassFilter filter = new ElementClassFilter(typeof(Sketch));
                                IList<ElementId> dependentRailIds = railing.GetDependentElements(filter);
                                var railingHostId = railing.HostId.ToString();
                                railingGeometry.Add("HostId", railingHostId);

                                var railingCrvLoop = new CurveLoop();
                                var railingList = railing.GetPath().ToList();
                                Debug.Write(railingList);
                                foreach (Curve railCrv in railingList)
                                {
                                    railingCrvLoop.Append(railCrv);
                                }
                                var railingCurveLoop = GetCurveLoopDescription(railingCrvLoop);
                                railingGeometry.Add("CurveLoop", railingCurveLoop);

                                // Property
                                JObject railingProperty = (JObject)addJarray["Property"];
                                var railingFlipped = railing.Flipped;
                                railingProperty.Add("Flipped", railingFlipped);

                                // Parameter
                                var railingParameters = railing.Parameters;
                                foreach (Parameter p in railingParameters)
                                {
                                    GetParameter(p, addJarray);
                                }
                            }
                            return addJarray;

                        case "OST_Windows":
                            // 24.3.24. 수정
                            JObject windowCommon = (JObject)addJarray["Common"];
                            JObject windowGeometry = (JObject)addJarray["Geometry"];
                            JObject windowProperty = (JObject)addJarray["Property"];
                            FamilyGetLayer(elem, addJarray, doc);
                            var window = elem as FamilyInstance;

                            // geometry
                            var windowHostId = window.Host.Id.Value;
                            var windowLocation = GetXYZDescription((window.Location as LocationPoint).Point);
                            windowGeometry.Add("HostId", windowHostId);
                            windowGeometry.Add("Location", windowLocation);
                            // property
                            var windowFlipFacing = window.FacingFlipped;
                            var windowFlipHand = window.HandFlipped;
                            // 24.7.15. 수정
                            var windowHeight = doc.GetElement(window.GetTypeId()).get_Parameter(BuiltInParameter.DOOR_HEIGHT).AsValueString();
                            var windowWidth = doc.GetElement(window.GetTypeId()).get_Parameter(BuiltInParameter.FURNITURE_WIDTH).AsValueString();
                            var winHeight = doc.GetElement(window.GetTypeId()).get_Parameter(BuiltInParameter.FURNITURE_HEIGHT).AsValueString();
                            var winWidth = doc.GetElement(window.GetTypeId()).get_Parameter(BuiltInParameter.WINDOW_WIDTH).AsValueString();
                            windowProperty.Add("FlipFacing", windowFlipFacing);
                            windowProperty.Add("FlipHand", windowFlipHand);
                            windowProperty.Add("Height", windowHeight);
                            windowProperty.Add("Width", windowWidth);
                            //built-in
                            var windowParameters = window.Parameters;
                            foreach (Parameter p in windowParameters)
                            {
                                GetParameter(p, addJarray);
                            }
                            return addJarray;

                        case "OST_Doors":
                            // 24.4.3. 수정
                            FamilyGetLayer(elem, addJarray, doc);
                            var door = elem as FamilyInstance;
                            // geometry
                            JObject doorGeometry = (JObject)addJarray["Geometry"];
                            var doorHostId = door.Host.Id.ToString();
                            var doorLocation = GetXYZDescription((door.Location as LocationPoint).Point);
                            doorGeometry.Add("HostId", doorHostId);
                            doorGeometry.Add("Location", doorLocation);
                            // property

                            JObject doorProperty = (JObject)addJarray["Property"];
                            var doorFlipFacing = door.FacingFlipped;
                            var doorFlipHand = door.HandFlipped;
                            var doorHeight = doc.GetElement(door.GetTypeId()).get_Parameter(BuiltInParameter.DOOR_HEIGHT).AsValueString();
                            var doorWidth = doc.GetElement(door.GetTypeId()).get_Parameter(BuiltInParameter.DOOR_WIDTH).AsValueString();
                            doorProperty.Add("FlipFacing", doorFlipFacing);
                            doorProperty.Add("FlipHand", doorFlipHand);
                            doorProperty.Add("Height", doorHeight);
                            doorProperty.Add("Width", doorWidth);
                            // parameters
                            var doorParameters = door.Parameters;
                            foreach (Parameter p in doorParameters)
                            {
                                GetParameter(p, addJarray);

                            }
                            return addJarray;

                        case "OST_Furniture":
                            FamilyGetLayer(elem, addJarray, doc);
                            var furniture = elem as FamilyInstance;
                            var furnitureLocation = GetXYZDescription((furniture.Location as LocationPoint).Point);
                            JObject furnitureGeometry = (JObject)addJarray["Geometry"];
                            furnitureGeometry.Add("Location", furnitureLocation);
                            var furnitureParameters = furniture.Parameters;
                            foreach (Parameter p in furnitureParameters)
                            {
                                GetParameter(p, addJarray);
                            }

                            return addJarray;

                        case "OST_Columns":
                            // 24.3.24. 수정
                            FamilyGetLayer(elem, addJarray, doc);
                            var column = elem as FamilyInstance;
                            // geometry
                            var columnLocation = GetXYZDescription((column.Location as LocationPoint).Point);
                            JObject columnGeometry = (JObject)addJarray["Geometry"];
                            columnGeometry.Add("Location", columnLocation);
                            // property
                            // 24.3.24. 추가작업
                            // AsDouble 말고 AsValueString으로 받아옴
                            var columnWidth = doc.GetElement(column.GetTypeId()).GetParameters("Depth").First().AsValueString();
                            var columnDepth = doc.GetElement(column.GetTypeId()).GetParameters("Width").First().AsValueString();
                            var columnProperty = (JObject)addJarray["Property"];
                            columnProperty.Add("width", columnWidth);
                            columnProperty.Add("Depth", columnDepth);
                            // parameter 
                            var columnParameters = column.Parameters;
                            foreach (Parameter p in columnParameters)
                            {
                                GetParameter(p, addJarray);

                            }

                            return addJarray;

                        case "OST_StructuralColumns":
                            // 24.4.5. 수정
                            FamilyGetLayer(elem, addJarray, doc);
                            var structuralcolumn = elem as FamilyInstance;
                            bool structuralcolumnIsCurve = false;
                            JObject structuralcolumnLocation;
                            if (elem.Location as LocationCurve != null)
                            {
                                structuralcolumnIsCurve = true;
                            }

                            if (structuralcolumnIsCurve == true)
                            {
                                structuralcolumnLocation = GetCurveDescription((structuralcolumn.Location as LocationCurve).Curve);
                            }
                            else
                            {
                                structuralcolumnLocation = GetXYZDescription((structuralcolumn.Location as LocationPoint).Point);
                            }
                            // Geometry
                            JObject structuralColumnGeometry = (JObject)addJarray["Geometry"];
                            structuralColumnGeometry.Add("Location", structuralcolumnLocation);

                            var structuralcolumnHeight = doc.GetElement(structuralcolumn.GetTypeId()).GetParameters("Ht").First().AsDouble();
                            var structuralcolumnThickness = doc.GetElement(structuralcolumn.GetTypeId()).GetParameters("t").First().AsDouble();

                            // property
                            JObject structuralColumnProperty = (JObject)addJarray["Property"];
                            structuralColumnProperty.Add("Height", structuralcolumnHeight);
                            structuralColumnProperty.Add("Thickness", structuralcolumnThickness);

                            //parameter
                            var structuralColumnParameters = structuralcolumn.Parameters;
                            foreach (Parameter p in structuralColumnParameters)
                            {
                                GetParameter(p, addJarray);
                            }

                            return addJarray;

                        case "OST_CurtainWallMullions":

                            var curtainmullion = elem as Mullion;
                            JObject mullionGeometry = (JObject)addJarray["Geometry"];
                            JObject mullionProperty = (JObject)addJarray["Property"];
                            var curtainmullionHostId = curtainmullion.Host.Id.ToString();
                            var curtainmullionCurve = GetCurveDescription(curtainmullion.LocationCurve);
                            mullionGeometry.Add("HostId", curtainmullionHostId);
                            mullionGeometry.Add("Curve", curtainmullionCurve);
                            var curtainmullionLocation = GetXYZDescription((curtainmullion.Location as LocationPoint).Point);
                            mullionProperty.Add("Location", curtainmullionLocation);
                            var curtainmullionParameters = curtainmullion.Parameters;
                            foreach (Parameter p in curtainmullionParameters)
                            {
                                GetParameter(p, addJarray);
                            }

                            return addJarray;

                        case "OST_StructuralFraming":
                            // 24.4.5. 수정
                            FamilyGetLayer(elem, addJarray, doc);
                            var structuralframing = elem as FamilyInstance;
                            // geometry
                            var structuralframingLocationCurve = GetCurveDescription((structuralframing.Location as LocationCurve).Curve);
                            JObject beamGeometry = (JObject)addJarray["Geometry"];
                            beamGeometry.Add("LocationCurve", structuralframingLocationCurve);
                            // property
                            //var structuralframingHeight = doc.GetElement(structuralframing.Symbol.Id).GetParameters("Height").First().AsDouble();
                            //var structuralframingWidth = doc.GetElement(structuralframing.Symbol.Id).GetParameters("Width").First().AsDouble();
                            var structuralframingH = doc.GetElement(structuralframing.Symbol.Id).get_Parameter(BuiltInParameter.STRUCTURAL_SECTION_COMMON_HEIGHT).AsValueString();
                            var structuralframingW = doc.GetElement(structuralframing.Symbol.Id).get_Parameter(BuiltInParameter.STRUCTURAL_SECTION_COMMON_WIDTH).AsValueString();
                            JObject beamProperty = (JObject)addJarray["Property"];
                            beamProperty.Add("Height", structuralframingH);
                            beamProperty.Add("Width", structuralframingW);
                            // parameter
                            var beamParameters = structuralframing.Parameters;
                            foreach (Parameter p in beamParameters)
                            {
                                GetParameter(p, addJarray);
                            }

                            return addJarray;

                        case "OST_Cornices":
                            //2024.07.01 wall sweep이 바뀔 때, wall이 modify됨

                            JObject wallsweepGeometry = (JObject)addJarray["Geometry"];
                            JObject wallsweepProperty = (JObject)addJarray["Property"];
                            var wallsweep = elem as WallSweep;
                            // instance geometry
                            var wallsweepWallId = GetIdListDescription(wallsweep.GetHostIds().ToList());
                            // properties
                            var wallsweepCutsWall = wallsweep.GetWallSweepInfo().CutsWall;
                            var wallsweepDefaultSetback = wallsweep.GetWallSweepInfo().DefaultSetback;
                            var wallsweepDistance = wallsweep.GetWallSweepInfo().Distance;
                            int wallsweepDistanceMeasuredFrom;
                            if (wallsweep.GetWallSweepInfo().DistanceMeasuredFrom == DistanceMeasuredFrom.Base)
                            {
                                wallsweepDistanceMeasuredFrom = 0;
                            }
                            else
                            {
                                wallsweepDistanceMeasuredFrom = 1;
                            }
                            var wallsweepId = wallsweep.GetWallSweepInfo().Id;
                            var wallsweepIsCutByInserts = wallsweep.GetWallSweepInfo().IsCutByInserts;
                            var wallsweepIsProfileFlipped = wallsweep.GetWallSweepInfo().IsProfileFlipped;
                            var wallsweepIsVertical = wallsweep.GetWallSweepInfo().IsVertical;
                            var wallsweepMaterialId = wallsweep.GetWallSweepInfo().MaterialId.ToString();
                            var wallsweepProfileId = wallsweep.GetWallSweepInfo().ProfileId.ToString();
                            var wallsweepWallOffset = wallsweep.GetWallSweepInfo().WallOffset;
                            int wallsweepWallSide;
                            if (wallsweep.GetWallSweepInfo().WallSide == WallSide.Exterior)
                            {
                                wallsweepWallSide = 0;
                            }
                            else
                            {
                                wallsweepWallSide = 1;

                            }
                            int wallsweepWallSweepOrientation;
                            if (wallsweep.GetWallSweepInfo().WallSweepOrientation == WallSweepOrientation.Horizontal)
                            {
                                wallsweepWallSweepOrientation = 0;
                            }
                            else
                            {
                                wallsweepWallSweepOrientation = 1;
                            }
                            int wallsweepWallSweepType;
                            if (wallsweep.GetWallSweepInfo().WallSweepType == WallSweepType.Sweep)
                            {
                                wallsweepWallSweepType = 0;
                            }
                            else
                            {
                                wallsweepWallSweepType = 1;
                            }

                            wallsweepGeometry.Add("WallId", wallsweepWallId);

                            wallsweepProperty.Add("CutsWall", wallsweepCutsWall);
                            wallsweepProperty.Add("DefaultSetback", wallsweepDefaultSetback);
                            wallsweepProperty.Add("Distance", wallsweepDistance);
                            wallsweepProperty.Add("DistanceMeasuredFrom", wallsweepDistanceMeasuredFrom);
                            wallsweepProperty.Add("Id", wallsweepId);
                            wallsweepProperty.Add("IsCutByInserts", wallsweepIsCutByInserts);
                            wallsweepProperty.Add("IsProfileFlipped", wallsweepIsProfileFlipped);
                            wallsweepProperty.Add("IsVertical", wallsweepIsVertical);
                            wallsweepProperty.Add("MaterialId", wallsweepMaterialId);
                            wallsweepProperty.Add("ProfileId", wallsweepProfileId);
                            wallsweepProperty.Add("WallOffset", wallsweepWallOffset);
                            wallsweepProperty.Add("WallSide", wallsweepWallSide);
                            wallsweepProperty.Add("WallSweepOrientation", wallsweepWallSweepOrientation);
                            wallsweepProperty.Add("WallSweepType", wallsweepWallSweepType);

                            var wallsweepParameters = wallsweep.Parameters;
                            foreach (Parameter p in wallsweepParameters)
                            {
                                GetParameter(p, addJarray);
                            }
                            return addJarray;

                        case "OST_Reveals":
                            //2024.07.01 reveals가 바뀔 때, wall이 modify됨

                            JObject revealsGeometry = (JObject)addJarray["Geometry"];
                            JObject revealsProperty = (JObject)addJarray["Property"];

                            var reveals = elem as WallSweep;

                            // Instance Geometry
                            // 굳이 루프문을 도는 이유가 있나??
                            var revealsWallId = GetIdListDescription(reveals.GetHostIds().ToList());

                            var revealsCutsWall = reveals.GetWallSweepInfo().CutsWall;
                            var revealsDefaultSetback = reveals.GetWallSweepInfo().DefaultSetback;
                            var revealsDistance = reveals.GetWallSweepInfo().Distance;
                            int revealsDistanceMeasureFrom;
                            if (reveals.GetWallSweepInfo().DistanceMeasuredFrom == DistanceMeasuredFrom.Base)
                            {
                                revealsDistanceMeasureFrom = 0;
                            }
                            else
                            {
                                revealsDistanceMeasureFrom = 1;
                            }
                            var revealsId = reveals.GetWallSweepInfo().Id;
                            var revealsIsCutByInserts = reveals.GetWallSweepInfo().IsCutByInserts;
                            var revealsIsProfileFlipped = reveals.GetWallSweepInfo().IsProfileFlipped;
                            var revealsIsVertical = reveals.GetWallSweepInfo().IsVertical;
                            var revealsMaterialId = reveals.GetWallSweepInfo().MaterialId.ToString();
                            var revealsProfileId = reveals.GetWallSweepInfo().ProfileId.ToString();
                            var revealsWallOffset = reveals.GetWallSweepInfo().WallOffset;
                            int revealsWallSide;
                            if (reveals.GetWallSweepInfo().WallSide == WallSide.Exterior)
                            {
                                revealsWallSide = 0;
                            }
                            else
                            {
                                revealsWallSide = 1;
                            }
                            int revealsWallSweepOrientation;
                            if (reveals.GetWallSweepInfo().WallSweepOrientation == WallSweepOrientation.Horizontal)
                            {
                                revealsWallSweepOrientation = 0;
                            }
                            else
                            {
                                revealsWallSweepOrientation = 1;
                            }
                            int revealsWallSweepType;
                            if (reveals.GetWallSweepInfo().WallSweepType == WallSweepType.Sweep)
                            {
                                revealsWallSweepType = 0;
                            }
                            else
                            {
                                revealsWallSweepType = 1;
                            }

                            revealsGeometry.Add("WallId", revealsWallId);

                            revealsProperty.Add("CutsWall", revealsCutsWall);
                            revealsProperty.Add("DefaultSetback", revealsDefaultSetback);
                            revealsProperty.Add("Distance", revealsDistance);
                            revealsProperty.Add("DistanceMeasuredFrom", revealsDistanceMeasureFrom);
                            revealsProperty.Add("Id", revealsId);
                            revealsProperty.Add("IsCutByInserts", revealsIsCutByInserts);
                            revealsProperty.Add("IsProfileFlipped", revealsIsProfileFlipped);
                            revealsProperty.Add("IsVertical", revealsIsVertical);
                            revealsProperty.Add("MaterialId", revealsMaterialId);
                            revealsProperty.Add("ProfileId", revealsProfileId);
                            revealsProperty.Add("WallOffset", revealsWallOffset);
                            revealsProperty.Add("WallSide", revealsWallSide);
                            revealsProperty.Add("WallSweepOrientation", revealsWallSweepOrientation);
                            revealsProperty.Add("WallSweepType", revealsWallSweepType);

                            var revealsParameters = reveals.Parameters;
                            foreach (Parameter p in revealsParameters)
                            {
                                GetParameter(p, addJarray);

                            }

                            return addJarray;

                        case "OST_StructuralFoundation":

                            if (elem.GetType().Name == "FamilyInstance")
                            {
                                cat = "Structural Foundations: Isolated";
                                FamilyGetLayer(elem, addJarray, doc);
                                JObject isolatedfoundationGeometry = (JObject)addJarray["Geometry"];
                                var isolatedfoundation = elem as FamilyInstance;
                                //geometry
                                var isolatedfoundationLocation = GetXYZDescription((isolatedfoundation.Location as LocationPoint).Point);
                                isolatedfoundationGeometry.Add("Location", isolatedfoundationLocation);
                                //builtin parameter
                                var isolatedfoundationParameters = isolatedfoundation.Parameters;
                                foreach (Parameter p in isolatedfoundationParameters)
                                {
                                    GetParameter(p, addJarray);
                                }
                            }
                            else if (elem.GetType().Name == "WallFoundation")
                            {
                                cat = "Structural Foundations: Wall";
                                JObject wallfoundationGeometry = (JObject)addJarray["Geometry"];
                                var wallfoundation = elem as WallFoundation;
                                //geometry
                                var wallfoundationWallId = wallfoundation.WallId.ToString();
                                wallfoundationGeometry.Add("WallId", wallfoundationWallId);
                                //builtin parameter
                                var wallfoundationParameters = wallfoundation.Parameters;
                                foreach (Parameter p in wallfoundationParameters)
                                {
                                    GetParameter(p, addJarray);
                                }
                            }
                            else if (elem.GetType().Name == "Floor")
                            {
                                cat = "Structural Foundations: Slab";
                                GetLayer(elem, addJarray, doc);

                                var slabfoundation = elem as Floor;
                                var slabfoundationCheckAdded = slabfoundation.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsValueString();
                                if (slabfoundationCheckAdded == null)
                                {
                                    if (addedElementList.Contains(eid))
                                    {
                                        return null;
                                    }
                                }
                                else
                                {
                                    common["ElementCategory"] = cat;
                                    if (addedElementList.Contains(eid))
                                    {
                                        common["CommandType"] = "ADDED";
                                        addedElementList.Remove(eid);
                                    }
                                    JObject slabfoundationGeometry = (JObject)addJarray["Geometry"];


                                    var slabfoundationSketch = (doc.GetElement(slabfoundation.SketchId) as Sketch);
                                    var profileEmpty = slabfoundationSketch.Profile.IsEmpty;
                                    var slabfoundationEIds = slabfoundationSketch.GetAllElements();
                                    JObject slabfoundationSlopeArrow = new JObject();
                                    string slabfoundationSlope = null;
                                    JObject slabfoundationSpanDirection = new JObject();
                                    foreach (ElementId slabfoundationEId in slabfoundationEIds)
                                    {
                                        if (doc.GetElement(slabfoundationEId).Name == "Slope Arrow" || doc.GetElement(slabfoundationEId).Name == "경사 화살표")
                                        {
                                            var slabfoundationcrv = (doc.GetElement(slabfoundationEId) as CurveElement).GeometryCurve;
                                            slabfoundationSlopeArrow = GetCurveDescription(slabfoundationcrv);
                                            List<Parameter> slabfoundationParams = (doc.GetElement(slabfoundationEId) as ModelLine).GetOrderedParameters().ToList();
                                            foreach (var param in slabfoundationParams)
                                            {
                                                if (param.Definition.Name == "Slope" || param.Definition.Name == "경사")
                                                {
                                                    slabfoundationSlope = param.AsValueString();
                                                }
                                            }
                                        }
                                        if (doc.GetElement(slabfoundationEId).Name == "Span Direction Edges" || doc.GetElement(slabfoundationEId).Name == "스팬 방향 모서리")
                                        {
                                            var slabfoundationcrv = (doc.GetElement(slabfoundationEId) as CurveElement).GeometryCurve;
                                            slabfoundationSpanDirection = GetCurveDescription(slabfoundationcrv);
                                        }
                                    }
                                    var slabfoundationProfile = GetProfileDescription(slabfoundationSketch);
                                    slabfoundationGeometry.Add("Profile", slabfoundationProfile);
                                    if (slabfoundationSlopeArrow.HasValues == false)
                                    {
                                        slabfoundationGeometry.Add("Slope Arrow", "None");
                                    }
                                    else
                                    {
                                        slabfoundationGeometry.Add("Slope Arrow", slabfoundationSlopeArrow);
                                    }
                                    if (slabfoundationSpanDirection.HasValues == false)
                                    {
                                        slabfoundationGeometry.Add("Span Direction", "None");
                                    }
                                    else
                                    {
                                        slabfoundationGeometry.Add("Span Direction", slabfoundationSpanDirection);
                                    }
                                    // floor Parameter
                                    var slabfoundationParameters = slabfoundation.Parameters;
                                    foreach (Parameter p in slabfoundationParameters)
                                    {
                                        GetParameter(p, addJarray);
                                    }
                                }
                            }
                            return addJarray;
                    }
                }
                return null;
            }
        }
        #endregion
        #region Curve Log
        //CurveArrArray, CurveLoop, Curve Classes

        public JObject GetProfileDescription( Sketch sketch )
        {
            CurveArrArray crvArrArr = sketch.Profile;
            JArray profileJArray = new JArray();
            JObject profileJObject = new JObject();
            profileJObject["profile"] = profileJArray;

            foreach (CurveArray CrvArr in crvArrArr)
            {
                foreach (Curve Crv in CrvArr)
                {
                    JObject description = GetCurveDescription(Crv);
                    profileJArray.Add(description);
                }
            }
            return profileJObject;
        }

        public JObject GetCurveLoopDescription( CurveLoop curveLoop )
        {
            List<Curve> crvList = curveLoop.ToList();

            JArray jarray = new JArray();
            JObject jobject = new JObject();
            jobject["curveLoop"] = jarray;

            foreach (Curve crv in crvList)
            {
                JObject description = GetCurveDescription(crv);
                jarray.Add(description);
            }
            return jobject;
        }

        public JObject GetCurveListDescription( List<Curve> crvList )
        {
            JArray jarray = new JArray();
            JObject jobject = new JObject();
            jobject["curveList"] = jarray;

            foreach (Curve crv in crvList)
            {
                JObject description = GetCurveDescription(crv);
                jarray.Add(description);
            }
            return jobject;
        }

        public dynamic GetCurveDescription( Curve crv )
        {
            // 24.4.19. 작업

            JObject JCurve = new JObject();
            dynamic description;
            string typ = crv.GetType().Name;

            switch (typ)
            {
                case "Line":
                    // = new LineDescription() { Type = typ, StartPoint = "\"" + crv.GetEndPoint(0).ToString().Replace(" ", String.Empty) + "\"", EndPoint = "\"" + crv.GetEndPoint(1).ToString().Replace(" ", String.Empty) + "\"" };
                    JCurve.Add("Type", typ);
                    JCurve.Add("endPoints", new JArray(
                        new JObject(
                            new JProperty("X", crv.GetEndPoint(0).X),
                            new JProperty("Y", crv.GetEndPoint(0).Y),
                            new JProperty("Z", crv.GetEndPoint(0).Z)
                            ),
                        new JObject(
                            new JProperty("X", crv.GetEndPoint(1).X),
                            new JProperty("Y", crv.GetEndPoint(1).Y),
                            new JProperty("Z", crv.GetEndPoint(1).Z)
                            )
                        ));
                    break;

                case "Arc":

                    var arc = crv as Arc;
                    var arcCen = arc.Center;
                    var arcNorm = arc.Normal;
                    var rad = arc.Radius;
                    var arcXAxis = arc.XDirection;
                    var arcYAxis = arc.YDirection;
                    var plane = Plane.CreateByNormalAndOrigin(arcNorm, arcCen);
                    var startDir = (arc.GetEndPoint(0) - arcCen).Normalize();
                    var endDir = (arc.GetEndPoint(1) - arcCen).Normalize();
                    var startAngle = arcXAxis.AngleOnPlaneTo(startDir, arcNorm);
                    var endAngle = arcXAxis.AngleOnPlaneTo(endDir, arcNorm);

                    JCurve.Add("Type", typ);
                    JCurve.Add("center", new JObject(
                        new JProperty("X", arcCen.X),
                        new JProperty("Y", arcCen.Y),
                        new JProperty("Z", arcCen.Z)
                        ));
                    JCurve.Add("radius", rad);
                    JCurve.Add("startAngle", startAngle);
                    JCurve.Add("endAngle", endAngle);
                    JCurve.Add("xAxis", new JObject(
                        new JProperty("X", arcXAxis.X),
                        new JProperty("Y", arcXAxis.Y),
                        new JProperty("Z", arcXAxis.Z)
                        ));
                    JCurve.Add("yAxis", new JObject(
                        new JProperty("X", arcYAxis.X),
                        new JProperty("Y", arcYAxis.Y),
                        new JProperty("Z", arcYAxis.Z)
                        ));

                    //description = new ArcDescription() { Type = typ, Center = "\"" + arcCen.ToString().Replace(" ", String.Empty) + "\"", Radius = rad, StartAngle = startAngle, EndAngle = endAngle, xAxis = "\"" + arcXAxis.ToString().Replace(" ", String.Empty) + "\"", yAxis = "\"" + arcYAxis.ToString().Replace(" ", String.Empty) + "\"" };

                    break;
                case "Ellipse":
                    // 24.4.19. 작업 중
                    var ellip = crv as Ellipse;
                    var cen = ellip.Center;
                    var xRad = ellip.RadiusX;
                    var yRad = ellip.RadiusY;
                    var xAxis = ellip.XDirection;
                    var yAxis = ellip.YDirection;
                    var startParam = ellip.GetEndParameter(0);
                    var endParam = ellip.GetEndParameter(1);
                    JCurve.Add("type", typ);
                    JCurve.Add("center", new JObject(
                        new JProperty("X", cen.X),
                        new JProperty("Y", cen.Y),
                        new JProperty("Z", cen.Z)
                        ));
                    JCurve.Add("radiusX", xRad);
                    JCurve.Add("radiusY", yRad);
                    JCurve.Add("xDirection", new JObject(
                        new JProperty("X", xAxis.X),
                        new JProperty("Y", xAxis.Y),
                        new JProperty("Z", xAxis.Z)
                        ));
                    JCurve.Add("yDirection", new JObject(
                        new JProperty("X", yAxis.X),
                        new JProperty("Y", yAxis.Y),
                        new JProperty("Z", yAxis.Z)
                        ));
                    JCurve.Add("startParameter", startParam);
                    JCurve.Add("endParameter", endParam);

                    break;

                case "HermiteSpline":

                    var herSpl = crv as HermiteSpline;
                    var contPts = herSpl.ControlPoints;
                    Int32 tangentCount = (herSpl.Tangents.Count - 1);
                    var startTangents = herSpl.Tangents[0].Normalize();
                    var endTangents = herSpl.Tangents[tangentCount].Normalize();
                    JArray hJarray = new JArray();
                    for (int i = 0 ; i < contPts.Count ; i++)
                    {
                        XYZ pt = contPts[i];
                        hJarray.Add(new JObject(
                            new JProperty("X", pt.X),
                            new JProperty("y", pt.Y),
                            new JProperty("Z", pt.Z)
                            ));
                    }
                    var periodic = herSpl.IsPeriodic;
                    JCurve.Add("controlPoints", hJarray);
                    JCurve.Add("isPeriodic", periodic);
                    JCurve.Add("startTagents", new JObject(
                        new JProperty("X", startTangents.X),
                        new JProperty("Y", startTangents.Y),
                        new JProperty("Z", startTangents.Z)
                        ));
                    JCurve.Add("endTagents", new JObject(
                        new JProperty("X", endTangents.X),
                        new JProperty("Y", endTangents.Y),
                        new JProperty("Z", endTangents.Z)
                        ));

                    break;


                case "CylindricalHelix":

                    var cylinHelix = crv as CylindricalHelix;
                    var basePoint = cylinHelix.BasePoint;
                    var radius = cylinHelix.Radius;
                    var xVector = cylinHelix.XVector;
                    var zVector = cylinHelix.ZVector;
                    var pitch = cylinHelix.Pitch;

                    var cylPlane = Plane.CreateByNormalAndOrigin(zVector, basePoint);
                    var cylStartDir = (cylinHelix.GetEndPoint(0) - basePoint).Normalize();
                    var cylEndDir = (cylinHelix.GetEndPoint(1) - basePoint).Normalize();

                    var cylStartAngle = cylStartDir.AngleOnPlaneTo(xVector, zVector);
                    var cylEndAngle = cylEndDir.AngleOnPlaneTo(xVector, zVector);

                    //description = new CylindricalHelixDescription() { Type = typ, BasePoint = "\"" + basePoint.ToString().Replace(" ", String.Empty) + "\"", Radius = radius, xVector = "\"" + xVector.ToString().Replace(" ", String.Empty) + "\"", zVector = "\"" + zVector.ToString().Replace(" ", String.Empty) + "\"", Pitch = pitch, StartAngle = cylStartAngle, EndAngle = cylEndAngle };

                    break;

                case "NurbSpline":

                    var nurbsSpl = crv as NurbSpline;
                    var degree = nurbsSpl.Degree;

                    string knots = "\"";
                    for (int i = 0 ; i < nurbsSpl.Knots.OfType<double>().ToList().Count ; i++)
                    {
                        double knot = nurbsSpl.Knots.OfType<double>().ToList()[i];
                        if (i != 0)
                        {
                            knots += ";";
                        }
                        knots += knot;
                    }
                    knots += "\"";

                    string nurbsCtrlPts = "\"";
                    for (int i = 0 ; i < nurbsSpl.CtrlPoints.Count ; i++)
                    {
                        XYZ pt = nurbsSpl.CtrlPoints[i];
                        if (i != 0)
                        {
                            nurbsCtrlPts += ";";
                        }
                        nurbsCtrlPts += pt.ToString().Replace(" ", String.Empty);
                    }
                    nurbsCtrlPts += "\"";

                    string weights = "\"";
                    for (int i = 0 ; i < nurbsSpl.Weights.OfType<double>().ToList().Count ; i++)
                    {
                        double weight = nurbsSpl.Weights.OfType<double>().ToList()[i];
                        if (i != 0)
                        {
                            weights += ";";
                        }
                        weights += weight;
                    }
                    weights += "\"";


                    //description = new NurbSplineDescription() { Type = typ, Degree = degree, Knots = knots, ControlPoints = nurbsCtrlPts, Weights = weights };

                    break;

                default:
                    description = null;
                    break;

            }
            return JCurve;
            //return description;
        }
        public JObject GetPlaneDescription( Plane plane )
        {
            JObject job = new JObject();
            job.Add("planeOrigin", new JObject(
                new JProperty("X", plane.Origin.X),
                new JProperty("Y", plane.Origin.Z),
                new JProperty("Z", plane.Origin.Y)
                ));
            job.Add("planeXVec", new JObject(
                new JProperty("X", plane.XVec.X),
                new JProperty("Y", plane.XVec.Z),
                new JProperty("Z", plane.XVec.Y)
                ));
            job.Add("planeYVec", new JObject(
                new JProperty("X", plane.YVec.X),
                new JProperty("Y", plane.YVec.Z),
                new JProperty("Z", plane.YVec.Y)
                ));

            return job;
        }
        public JObject GetXYZDescription( XYZ xyz )
        {

            JObject jobject = new JObject(
        new JProperty("X", xyz.X),
        new JProperty("Y", xyz.Y),
        new JProperty("Z", xyz.Z)
        );

            return jobject;
        }

        public JArray GetIdListDescription( List<ElementId> elementIds )
        {
            JArray jarray = new JArray();
            for (int i = 0 ; i < elementIds.Count ; i++)
            {
                jarray.Add(elementIds[i].Value);
            }

            return jarray;
        }

        //public string GetRooms( Element element )
        // {
        //    string roomString = "[";
        //    try
        //    {

        //    Document doc = element.Document;
        //    BoundingBoxXYZ element_bb = element.get_BoundingBox(null);
        //    Outline outline = new Outline(element_bb.Min, element_bb.Max);
        //    BoundingBoxIntersectsFilter bbfilter = new BoundingBoxIntersectsFilter(outline);
        //    List<Element> room_elems = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WherePasses(bbfilter).ToElements().ToList();

        //    for (int i = 0 ; i < room_elems.Count ; i++)
        //    {
        //        ElementId eid = room_elems[i].Id;
        //        if (i != 0)
        //        {
        //            roomString += ";";
        //        }
        //        roomString += eid.ToString();
        //    }
        //    roomString += "]";

        //    if (roomString == "[]")
        //    {
        //        roomString = null;
        //    }

        //    }
        //        catch(Exception ex)
        //    {
        //        Debug.WriteLine(ex);
        //    }
        //    return roomString;
        //}
        #endregion



    }
}