using Microsoft.Win32;
using Microsoft.Win32.TaskScheduler;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Formats.Asn1;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text.Json;
using System.Web;
using Task = System.Threading.Tasks.Task;

public class desktop_wallpaper_csharp
{
    private const string BingImageApiUrl = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US";
    private static readonly string StartupKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);

    static async Task Main(string[] args)
    {
        try
        {
            // 下载图片
            string imageUrl = await GetBingImageUrl();
            string path = await DownloadImage(imageUrl);

            // 设置图片为壁纸
            SetWallpaper(path);

            // 注册开机启动
            //RegisterStartup();
        }
        catch (System.Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    private static async Task<string> GetBingImageUrl()
    {
        using (HttpClient client = new HttpClient())
        {
            string json = await client.GetStringAsync(BingImageApiUrl);
            //Dictionary<string, object> result = JsonSerializer.Deserialize<Dictionary<string, object>>(json);
            dynamic result = JsonSerializer.Deserialize<JsonElement>(json);
            Console.WriteLine(result);
            //dynamic result = JObject.Parse(json);
            //dynamic result = Json.Decode(json);
            //dynamic result = Newtonsoft.Json.JsonConvert.DeserializeObject(json);
            //JavaScriptSerializer jss = new JavaScriptSerializer();
            //var result = jss.Deserialize<dynamic>(json);
            //dynamic result = jss.Deserialize<object>(json);
            //var reader = new JsonReader();
            //dynamic result = reader.Read(json);
            //var writer = new JsonWriter();
            //string json = writer.Write(output);
            //dynamic result = JsonHelper.FromJson<dynamic>(json);
            string url = "https://www.bing.com" + result.GetProperty("images")[0].GetProperty("url");
            return url;
        }
    }

    private static async Task<string> DownloadImage(string imageUrl)
    {
        using (HttpClient client = new HttpClient())
        {
            byte[] imageBytes = await client.GetByteArrayAsync(imageUrl);

            NameValueCollection parameters = HttpUtility.ParseQueryString(new System.Uri(imageUrl).Query);
            string id = parameters.Get("id");
            Console.WriteLine(id);
            string rf = parameters.Get("rf");
            Console.WriteLine(rf);
            // 获取文件后缀
            string ext = Path.GetExtension(rf);
            if (ext == null)
            {
                ext = ".jpg";
            }
            // Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
            // Directory.GetCurrentDirectory()
            Console.WriteLine(AppDomain.CurrentDomain.BaseDirectory);
            string WallpaperPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bing_wallpaper" + ext);

            File.WriteAllBytes(WallpaperPath, imageBytes);

            return WallpaperPath;
        }
    }

    private static void SetWallpaper(string filePath)
    {
        SystemParametersInfo(20, 0, filePath, 0x01 | 0x02);
    }

    private static void RegisterStartup()
    {
        using (RegistryKey key = Registry.CurrentUser.OpenSubKey(StartupKey, true))
        {
            key.SetValue("BingWallpaperApp", $"\"{System.Reflection.Assembly.GetExecutingAssembly().Location}\"");
        }
    }

    // 添加任务计划：系统启动、登录或解锁时触发
    public static void AddTaskSchedulerEntry(string taskName, string description, string appPath)
    {
        using (TaskService ts = new TaskService())
        {
            // Create a new task definition and assign properties
            TaskDefinition td = ts.NewTask();
            //td.Principal.UserId = "SYSTEM";
            td.Principal.UserId = WindowsIdentity.GetCurrent().Name;
            td.Principal.RunLevel = TaskRunLevel.Highest;
            td.Principal.LogonType = TaskLogonType.InteractiveToken;
            td.Settings.RunOnlyIfIdle = false;
            td.Settings.Enabled = true;
            td.Settings.Hidden = false;
            td.Settings.WakeToRun = true;
            td.Settings.AllowDemandStart = true;
            td.Settings.StartWhenAvailable = true;
            td.Settings.DisallowStartIfOnBatteries = false;
            td.Settings.DisallowStartOnRemoteAppSession = false;
            td.Settings.StopIfGoingOnBatteries = false;
            td.Settings.IdleSettings.StopOnIdleEnd = false;
            td.Settings.MultipleInstances = TaskInstancesPolicy.IgnoreNew;
            td.Settings.RestartCount = 5;
            td.Settings.RestartInterval = TimeSpan.FromSeconds(100);
            td.Settings.UseUnifiedSchedulingEngine = true;
            td.Settings.Priority = ProcessPriorityClass.High;
            td.Settings.ExecutionTimeLimit = TimeSpan.Zero;
            //td.Settings.ExecutionTimeLimit = new TimeSpan(0);
            //td.Settings.ExecutionTimeLimit = TimeSpan.FromSeconds(0);
            td.RegistrationInfo.Source = appPath;
            td.RegistrationInfo.Description = description;
            td.RegistrationInfo.Author = "Andrew Sampson";
            td.RegistrationInfo.Date = new DateTime();


            /*var trigger = new DailyTrigger
            {
                StartBoundary = DateTime.Today + TimeSpan.FromHours(hour) + TimeSpan.FromMinutes(min)
            };*/

            var lt = new LogonTrigger();
            lt.Delay = TimeSpan.FromSeconds(5);
            lt.Enabled = true;
            lt.StartBoundary = DateTime.Now;
            lt.UserId = WindowsIdentity.GetCurrent().Name;
            // Create a trigger for logon
            td.Triggers.Add(lt);
            //td.Triggers.Add(new RegistrationTrigger());

            // Create an action that will launch the program whenever the trigger fires
            td.Actions.Add(new ExecAction(appPath, null, null));
            //td.Actions.Add(new ExecAction(appPath, "-FromTaskManager"));

            // Register the task in the root folder of the task scheduler
            ts.RootFolder.RegisterTaskDefinition(taskName, td);
            //var path = Path.Combine("Lazurite", "Execute");
            //ts.RootFolder.RegisterTaskDefinition(path, td, TaskCreation.CreateOrUpdate, WindowsIdentity.GetCurrent().Name, logonType: TaskLogonType.InteractiveToken);

            //ts.FindTask(taskName).Run();
            ts.Dispose();
        }
    }

    public static void RemoveLogonTask(string taskName)
    {
        using (TaskService ts = new TaskService())
        {
            ts.FindTask(taskName).Stop();
            ts.RootFolder.DeleteTask(taskName, false);
            ts.Dispose();
        }
    }
}
