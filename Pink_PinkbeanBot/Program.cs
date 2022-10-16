using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Pink_PinkbeanBot
{
    internal class Program
    {
        internal class Message { }

        static string token = "Your SlackBot Token";
        static string[] channel = { "ChannelID" };

        static string[] Name = { "Name" };
        static string[] MemberID = { "MemberID" };

        static string[] ChatbotSheetName = new string[Name.Length];
        static string[] ChatbotSheetLT = new string[Name.Length];
        static string[] ChatbotSheetResult = new string[Name.Length];
        static string[] ChatbotUpdateTime = new string[Name.Length];

        static StringCollection Mention = new StringCollection();

        static protected ChromeDriverService _driverService = null;
        static protected ChromeOptions _options = null;
        static protected ChromeDriver _driver = null;

        static void Main()
        {
            Console.WriteLine("Bot Start");

            /*using (var client = new HttpClient())
            {
                var url = "https://api.nexon.co.kr/fifaonline4/v1.0/users/58c664aeb2eeb9f9b52f6916/maxdivision";
                client.DefaultRequestHeaders.Add("Authorization", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJYLUFwcC1SYXRlLUxpbWl0IjoiNTAwOjEwIiwiYWNjb3VudF9pZCI6IjE5MzAwNTYxNDIiLCJhdXRoX2lkIjoiMiIsImV4cCI6MTY3OTA0NjU5NCwiaWF0IjoxNjYzNDk0NTk0LCJuYmYiOjE2NjM0OTQ1OTQsInNlcnZpY2VfaWQiOiI0MzAwMTE0ODEiLCJ0b2tlbl90eXBlIjoiQWNjZXNzVG9rZW4ifQ.BqSyw46PpHa6ZCS8tur8UCi9oYt40-t2JnmXep7FEl8");
                var response = client.GetStringAsync(url);
                Console.WriteLine(response.Result);
            }*/

            #region Settings
            if (Properties.Settings.Default.Mention != null)
            {
                Mention = Properties.Settings.Default.Mention;
            }
            #endregion

            #region Selenium
            _driverService = ChromeDriverService.CreateDefaultService();
            _driverService.HideCommandPromptWindow = true;
            _options = new ChromeOptions();
            _options.AddArgument("disable-gpu");
            //_options.AddArgument("headless");
            _driver = new ChromeDriver(_driverService, _options);
            _driver.Manage().Window.Size = new Size(1000, 2000);
            _driver.Navigate().GoToUrl("https://loawa.com/");
            #endregion

            ThreadManage();
        }

        static void ThreadManage()
        {
            new Thread(Slack).Start();
            new Thread(Alarm).Start();
            new Thread(ChatbotSheets).Start();
        }

        #region COMMAND
        public static void Command(object contents)
        {
            try
            {
                string[] Message = contents as string[];

                if (Message[3] == "핑크빈 명령어")
                {
                    SendMessage(Message[4], "```※기본 명령어\r\n" +
                        "핑크빈 프로필\r\n\r\n" +
                        "※게임 전적 검색\r\n" +
                        "핑크빈 랭크 *닉네임*\r\n" +
                        "핑크빈 로아 *닉네임*\r\n\r\n" +
                        "※장기미배차\r\n" +
                        "핑크빈 장미배 *배달번호*\r\n" +
                        "장기미배차시트, 장미배시트\r\n" +
                        "핑크빈 알람호출\r\n\r\n" +
                        "※핑크빈 키우기\r\n" +
                        "준비중```");
                    Console.WriteLine(Message[1] + "님이 명령어 호출");
                }
                else if (Message[3] == "하이")
                {
                    SendMessage(Message[4], "안녕하세요.");
                    Console.WriteLine(Message[1] + "님이 인사 명령어 호출");
                }
                else if (Message[3] == "바이")
                {
                    SendMessage(Message[4], "안녕히가세요.");
                    Console.WriteLine(Message[1] + "님이 인사 명령어 호출");
                }
                else if (Message[3] == "장기미배차시트" || Message[3] == "장미배시트")
                {
                    SendMessage(Message[4], "*장기 미배차 매크로 처리현황 시트*\r\n" +
                        "https://docs.google.com/spreadsheets/d/16J8lrENJSbepH3Qooqzf4D2sYYOHJPQCmvtg3nOagco/edit?usp=sharing\r\n\r\n" +
                        "*장기 미배차 시트*\r\n" +
                        "https://docs.google.com/spreadsheets/d/1XA3LYMn2C9LkUcWrrvS9vN0Z3FX7q1pgXD6GymCj5uc/edit#gid=257145023");
                    Console.WriteLine(Message[1] + "님이 장기미배차 시트 명령어 호출");
                }
                else if (Message[3] == "핑크빈 프로필")
                {
                    #region 프로필
                    try
                    {
                        int membercount = Array.IndexOf(MemberID, Message[1]);
                        int arraycount = Array.IndexOf(ChatbotSheetName, Name[membercount]);

                        if (membercount != -1 || arraycount != -1)
                        {
                            SendMessage(Message[4], "*Lv 0 " + ChatbotSheetName[arraycount] + "님의 프로필*\n\n" +
                                "―――――――――――――――――――\n" +
                                "*리드타임*                          *처리건수*\n" +
                                ChatbotSheetLT[arraycount] + "                           " + ChatbotSheetResult[arraycount] + "\n" +
                                "갱신 시간 : " + ChatbotUpdateTime[arraycount], true);
                        }
                        else if (membercount == -1 || arraycount == -1)
                        {
                            SendMessage(Message[4], "관리자한테 승인된 사용자가 아니거나 시트 업데이트 중입니다.");
                        }
                        Console.WriteLine(Message[1] + "님이 프로필 명령어 호출");
                    }
                    catch (Exception)
                    {
                        SendMessage(Message[4], "관리자한테 승인된 사용자가 아니거나 시트 업데이트 중입니다.");
                    }
                    #endregion
                }
                else if (Message[3] == "핑크빈 알람호출")
                {
                    if (Mention.IndexOf(Message[1]) != -1)
                    {
                        Mention.Remove(Message[1]);
                        SendMessage(Message[4], "장기미배차 알람 멘션에 제거 완료 되었습니다.");
                    }
                    else
                    {
                        Mention.Add(Message[1]);
                        SendMessage(Message[4], "장기미배차 알람 멘션에 추가 완료 되었습니다.");
                    }
                    Properties.Settings.Default.Mention = Mention;
                    Properties.Settings.Default.Save();
                    Console.WriteLine(Message[1] + "님이 알람호출 명령어 호출");
                }
                else if (Message[3].Contains("핑크빈 랭크") == true)
                {
                    if (Message[3].Substring(0, 6) == "핑크빈 랭크")
                    {
                        #region 랭크
                        try
                        {
                            string SummonerName = Message[3].Replace("핑크빈 랭크 ", "");
                            LeagueOfLegends.summonerName = SummonerName;
                            Console.WriteLine(Message[1] + "님이 랭크 명령어 호출 : " + SummonerName);
                            if (SummonerName == "")
                            {
                                SendMessage(Message[4], "소환사명을 입력해주세요.");
                            }
                            else
                            {
                                string[] Champion = LeagueOfLegends.Champion();
                                string[] GameResult = LeagueOfLegends.GameResult();
                                string[] Kill = LeagueOfLegends.Kill();
                                string[] Death = LeagueOfLegends.Death();
                                string[] Assist = LeagueOfLegends.Assist();
                                string[] SummonerInfo = LeagueOfLegends.SummonerInfo();

                                double Percent = Convert.ToDouble(SummonerInfo[3]) / (Convert.ToDouble(SummonerInfo[3]) + Convert.ToDouble(SummonerInfo[4]));
                                Percent = Math.Truncate(Percent * 100) / 100;

                                if (SummonerInfo != null)
                                {
                                    SendMessage(Message[4], "**개인/2인/자유 랭크**\n" +
                                        SummonerInfo[0] + " " + SummonerInfo[2] + " " + SummonerInfo[5] + "LP\n" +
                                        SummonerInfo[3] + "승 " + SummonerInfo[4] + "패 " + Percent * 100 + "%" + "\n\n" +
                                        "**최근 전적**\n" +
                                        GameResult[0] + Champion[0] + " " + Kill[0] + "/" + Death[0] + "/" + Assist[0] + "\n" +
                                        GameResult[1] + Champion[1] + " " + Kill[1] + "/" + Death[1] + "/" + Assist[1] + "\n" +
                                        GameResult[2] + Champion[2] + " " + Kill[2] + "/" + Death[2] + "/" + Assist[2] + "\n" +
                                        GameResult[3] + Champion[3] + " " + Kill[3] + "/" + Death[3] + "/" + Assist[3] + "\n" +
                                        GameResult[4] + Champion[4] + " " + Kill[4] + "/" + Death[4] + "/" + Assist[3] + "\n" +
                                        GameResult[5] + Champion[5] + " " + Kill[5] + "/" + Death[5] + "/" + Assist[5] + "\n" +
                                        GameResult[6] + Champion[6] + " " + Kill[6] + "/" + Death[6] + "/" + Assist[6] + "\n" +
                                        GameResult[7] + Champion[7] + " " + Kill[7] + "/" + Death[7] + "/" + Assist[7] + "\n" +
                                        GameResult[8] + Champion[8] + " " + Kill[8] + "/" + Death[8] + "/" + Assist[8] + "\n" +
                                        GameResult[9] + Champion[9] + " " + Kill[9] + "/" + Death[9] + "/" + Assist[9]);
                                }
                                else
                                {
                                    SendMessage(Message[4], "**최근 전적**\n" +
                                        GameResult[0] + Champion[0] + " " + Kill[0] + "/" + Death[0] + "/" + Assist[0] + "\n" +
                                        GameResult[1] + Champion[1] + " " + Kill[1] + "/" + Death[1] + "/" + Assist[1] + "\n" +
                                        GameResult[2] + Champion[2] + " " + Kill[2] + "/" + Death[2] + "/" + Assist[2] + "\n" +
                                        GameResult[3] + Champion[3] + " " + Kill[3] + "/" + Death[3] + "/" + Assist[3] + "\n" +
                                        GameResult[4] + Champion[4] + " " + Kill[4] + "/" + Death[4] + "/" + Assist[3] + "\n" +
                                        GameResult[5] + Champion[5] + " " + Kill[5] + "/" + Death[5] + "/" + Assist[5] + "\n" +
                                        GameResult[6] + Champion[6] + " " + Kill[6] + "/" + Death[6] + "/" + Assist[6] + "\n" +
                                        GameResult[7] + Champion[7] + " " + Kill[7] + "/" + Death[7] + "/" + Assist[7] + "\n" +
                                        GameResult[8] + Champion[8] + " " + Kill[8] + "/" + Death[8] + "/" + Assist[8] + "\n" +
                                        GameResult[9] + Champion[9] + " " + Kill[9] + "/" + Death[9] + "/" + Assist[9]);
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            SendMessage(Message[4], "랭크 전적이 10전 미만이거나 찾을 수 없는 소환사 혹은 명령어 양식이 잘못 되었습니다.");
                            Console.WriteLine(e.Message);
                        }
                        #endregion
                    }
                }
                else if (Message[3].Contains("핑크빈 장미배") == true)
                {
                    if (Message[3].Substring(0, 7) == "핑크빈 장미배")
                    {
                        #region 장미배
                        try
                        {
                            string deliveryNum = Message[3].Replace("핑크빈 장미배 ", "");
                            Console.WriteLine(Message[1] + "님이 장미배 명령어 호출");

                            using (Socket svrSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp))
                            {
                                IPEndPoint endPoint = new IPEndPoint(IPAddress.Any, 9999);

                                svrSocket.Bind(endPoint);

                                svrSocket.Listen(100);
                                Socket cltSocket = svrSocket.Accept();

                                int membercount = Array.IndexOf(MemberID, Message[1]);

                                if (membercount != -1)
                                {
                                    Send(cltSocket, Name[membercount] + ":" + deliveryNum);
                                    SendMessage(Message[4], "장미배 추가 완료");
                                }
                                else
                                {
                                    SendMessage(Message[4], "인증된 사용자가 아닙니다.");
                                }
                                cltSocket.Close();
                            }
                        }
                        catch (Exception)
                        {
                            SendMessage(Message[4], "장미배 프로그램이 실행중이지 않습니다. 제작자에게 문의하여 주세요");
                        }
                        #endregion
                    }
                }
                else if (Message[3].Contains("핑크빈 로아") == true)
                {
                    if (Message[3].Substring(0, 6) == "핑크빈 로아")
                    {
                        #region 로아
                        try
                        {
                            string CharacterName = Message[3].Replace("핑크빈 로아 ", "");

                            bool characterinfocheck = File.Exists(@"C:\Bitnami\wampstack-8.0.12-0\apache2\htdocs\lostark\" + CharacterName + "_Characterinfo.png");

                            if (characterinfocheck == true)
                            {
                                File.Delete(@"C:\Bitnami\wampstack-8.0.12-0\apache2\htdocs\lostark\" + CharacterName + "_Characterinfo.png");
                            }

                            bool characteritemcheck = File.Exists(@"C:\Bitnami\wampstack-8.0.12-0\apache2\htdocs\lostark\" + CharacterName + "_Iteminfo.png");

                            if (characterinfocheck == true)
                            {
                                File.Delete(@"C:\Bitnami\wampstack-8.0.12-0\apache2\htdocs\lostark\" + CharacterName + "_Iteminfo.png");
                            }

                            _driver.Navigate().GoToUrl("https://loahae.com/profile/" + CharacterName);
                            Thread.Sleep(2000);

                            var characterinfo = _driver.FindElement(By.XPath("/html/body/div/div[2]/div[1]/div"));
                            var characterscr = NodeCaptureScreenShot(_driver, characterinfo);
                            characterscr.Save(String.Format(@"C:\Bitnami\wampstack-8.0.12-0\apache2\htdocs\lostark\" + CharacterName + "_Characterinfo.png", ImageFormat.Png));

                            _driver.Navigate().GoToUrl("https://loawa.com/char/" + CharacterName);
                            Thread.Sleep(2000);

                            var iteminfo = _driver.FindElement(By.XPath("/html/body/div[2]/main/article/div[7]/div/div[2]/div[2]/div/div/div[1]/div[1]/div"));
                            var itemscr = NodeCaptureScreenShot(_driver, iteminfo);
                            itemscr.Save(String.Format(@"C:\Bitnami\wampstack-8.0.12-0\apache2\htdocs\lostark\" + CharacterName + "_Iteminfo.png", ImageFormat.Png));

                            SendMessage(Message[4], CharacterName + "님의 로스트아크 캐릭터 정보입니다.", true, "http://175.125.94.208/lostark/" + CharacterName + "_Characterinfo.png");
                            SendMessage(Message[4], CharacterName + "님의 로스트아크 아이템 정보입니다.", true, "http://175.125.94.208/lostark/" + CharacterName + "_Iteminfo.png");
                            //SendMessage(Message[4], CharacterName + "님의 로스트아크 카드 정보입니다.", true, "http://175.125.94.208/Cardinfo.png");
                        }
                        catch (Exception e)
                        {
                            SendMessage(Message[4], "찾을 수 없는 캐릭터 혹은 명령어 양식이 잘못 되었습니다.");
                            Console.WriteLine(e.Message);
                        }
                        Console.WriteLine(Message[1] + "님이 로스트아크 명령어 호출");
                        #endregion
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        #endregion

        #region ALARM
        static void Alarm()
        {
            int beforecount = 0;

            while (true)
            {
                try
                {
                    string[] count = AlarmSheets();

                    if (Convert.ToInt32(count[2]) != beforecount)
                    {
                        var MentionStr = new StringBuilder();

                        foreach (var name in Mention)
                        {
                            MentionStr.Append("<@" + name + ">\n");
                        }

                        if (Convert.ToInt32(count[2]) == 0)
                        {
                            if (MentionStr.ToString() != "")
                            {
                                MentionStr.Clear();
                            }
                            MentionStr.Append("모든 문의건이 처리 되었습니다.");
                        }

                        if (MentionStr.ToString() == "")
                        {
                            MentionStr.Append("멘션 요청 인원이 없습니다.");
                        }

                        if (count[2].Length == 1)
                        {
                            SendMessage("C0416QDR5FW", "<https://docs.google.com/spreadsheets/d/1XA3LYMn2C9LkUcWrrvS9vN0Z3FX7q1pgXD6GymCj5uc/edit#gid=257145023|장기 미배차 알림> :최고농담곰-따봉농담곰: \n\n" +
                                "―――――――――――――――――――\n" +
                                "*미처리*                          *전체 처리*\n" +
                                count[2] + "                                   " + count[1] + "\n" +
                                "*멘션*\n" +
                                MentionStr, true);
                        }
                        else if (count[2].Length == 2)
                        {
                            SendMessage("C0416QDR5FW", "<https://docs.google.com/spreadsheets/d/1XA3LYMn2C9LkUcWrrvS9vN0Z3FX7q1pgXD6GymCj5uc/edit#gid=257145023|장기 미배차 알림> :최고농담곰-따봉농담곰: \n\n" +
                                "―――――――――――――――――――\n" +
                                "*미처리*                          *전체 처리*\n" +
                                count[2] + "                                " + count[1] + "\n" +
                                "*멘션*\n" +
                                MentionStr, true);
                        }
                        else
                        {
                            SendMessage("C0416QDR5FW", "<https://docs.google.com/spreadsheets/d/1XA3LYMn2C9LkUcWrrvS9vN0Z3FX7q1pgXD6GymCj5uc/edit#gid=257145023|장기 미배차 알림> :최고농담곰-따봉농담곰: \n\n" +
                                "―――――――――――――――――――\n" +
                                "*미처리*                          *전체 처리*\n" +
                                count[2] + "                              " + count[1] + "\n" +
                                "*멘션*\n" +
                                MentionStr, true);
                        }
                        beforecount = Convert.ToInt32(count[2]);
                    }
                    Thread.Sleep(100);
                }
                catch (Exception)
                {
                    Thread.Sleep(100);
                }
            }
        }
        #endregion

        #region FUNCTION
        public static Bitmap NodeCaptureScreenShot(IWebDriver driver, IWebElement elem)
        {
            Screenshot myScreenShot = ((ITakesScreenshot)driver).GetScreenshot();

            using (var screenBmp = new Bitmap(new MemoryStream(myScreenShot.AsByteArray)))
            {
                return screenBmp.Clone(new Rectangle(elem.Location, elem.Size), screenBmp.PixelFormat);
            }
        }
        #endregion

        #region SOCKET
        public static bool Send(Socket sock, String msg)
        {
            byte[] data = Encoding.Default.GetBytes(msg);
            int size = data.Length;

            byte[] data_size = new byte[4];
            data_size = BitConverter.GetBytes(size);
            sock.Send(data_size);

            sock.Send(data, 0, size, SocketFlags.None);
            return true;
        }

        public static bool Recieve(Socket sock, ref String msg)
        {
            byte[] data_size = new byte[4];
            sock.Receive(data_size, 0, 4, SocketFlags.None);
            int size = BitConverter.ToInt32(data_size, 0);
            byte[] data = new byte[size];
            sock.Receive(data, 0, size, SocketFlags.None);
            msg = Encoding.Default.GetString(data);
            return true;
        }
        #endregion

        #region API
        static void SendMessage(string channel, string text, bool attach = false, string imgurl = null)
        {
            var data = new NameValueCollection();
            data["token"] = token;
            data["channel"] = channel;
            data["as_user"] = "true";

            if (attach == false)
            {
                data["text"] = text;
            }
            if (attach == true)
            {
                if (imgurl == null)
                {
                    data["attachments"] = "[{\"fallback\":\"핑크빈 봇 알림\", \"color\":\"#0ABAB5\", \"text\":\"" + text + "\"}]";
                }
                else
                {
                    data["attachments"] = "[{\"fallback\":\"핑크빈 봇 알림\", \"color\":\"#0ABAB5\", \"text\":\"" + text + "\", \"image_url\":\"" + imgurl + "\"}]";
                }
            }

            var client = new WebClient();
            var response = client.UploadValues("https://slack.com/api/chat.postMessage", "POST", data);
            Encoding.UTF8.GetString(response);
        }

        public static void Slack()
        {
            string[] client_msg_id = new string[channel.Length];

            while (true)
            {
                try
                {
                    foreach (var ch in channel.Select((value, index) => (value, index)))
                    {
                        string[] contents = new string[5];
                        var data = new NameValueCollection();
                        string result = null;
                        Uri url = new Uri("https://slack.com/api/conversations.history");

                        data["token"] = token;
                        data["channel"] = ch.value;
                        data["limit"] = "1";

                        var client = new WebClient();
                        var response = client.UploadValues(url, "POST", data);
                        string responseInString = Encoding.UTF8.GetString(response);

                        JObject obj = JObject.Parse(responseInString);

                        try
                        {
                            contents[0] = obj["messages"][0]["bot_id"].ToString();

                            for (int i = 0; i < contents.Length; i++)
                            {
                                contents[i] = "봇 메세지입니다.";
                            }
                            result = contents[0];
                        }
                        catch (NullReferenceException)
                        {
                            contents[0] = obj["messages"][0]["client_msg_id"].ToString();

                            if (contents[0] != client_msg_id[ch.index])
                            {
                                contents[1] = obj["messages"][0]["user"].ToString();
                                contents[2] = obj["messages"][0]["type"].ToString();
                                contents[3] = obj["messages"][0]["text"].ToString();
                                contents[4] = ch.value;
                                client_msg_id[ch.index] = contents[0];
                                result = contents[3];
                                //Console.WriteLine(contents[1], contents[2], contents[3], contents[4]);
                                Thread command = new Thread(new ParameterizedThreadStart(Command));
                                command.Start(contents);
                            }
                            else if (contents[0] == client_msg_id[ch.index])
                            {
                                for (int i = 0; i < contents.Length; i++)
                                {
                                    contents[i] = "중복 메세지입니다.";
                                }
                                result = contents[0];
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }

        #region SHEETS
        static string[] AlarmSheets()
        {
            string[] str = new string[1];

            try
            {
                string[] Scopes = { SheetsService.Scope.Spreadsheets };
                var ApplicationName = "Google Sheets API";

                UserCredential credential;

                using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                }

                // API 서비스 생성
                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                var spreadsheetId = "1XigD8gPfyhpqZPjTtpW3pBKKzsUvG-LOCRChn1TTRO8";
                var range = "Dash";

                var request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                ValueRange response = request.Execute();
                IList<IList<Object>> values = response.Values;

                str = new string[values.Count];

                if (values != null && values.Count > 0)
                {
                    foreach (var row in values.Select((value, index) => (value, index)))
                    {
                        str[row.index] = row.value[0].ToString();
                    }
                }
                else
                {
                    str[0] = "0";
                    str[1] = "0";
                    str[2] = "0";
                }

                return str;
            }
            catch (Exception)
            {
                str[0] = "0";
                str[1] = "0";
                str[2] = "0";
                return str;
            }
        }

        static void ChatbotSheets()
        {
            while (true)
            {
                try
                {
                    string[] Scopes = { SheetsService.Scope.Spreadsheets };
                    var ApplicationName = "Google Sheets API";

                    UserCredential credential;

                    using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                    {
                        string credPath = "token.json";
                        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    }

                    // API 서비스 생성
                    var service = new SheetsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName,
                    });

                    var spreadsheetId = "1XigD8gPfyhpqZPjTtpW3pBKKzsUvG-LOCRChn1TTRO8";
                    var range = "Chatbot";

                    var request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                    ValueRange response = request.Execute();
                    IList<IList<Object>> values = response.Values;

                    if (values.Count > 0)
                    {
                        foreach (var row in values.Select((value, index) => (value, index)))
                        {
                            try
                            {
                                if (ChatbotSheetLT[row.index] != row.value[1].ToString() || ChatbotSheetResult[row.index] != row.value[2].ToString())
                                {
                                    ChatbotSheetName[row.index] = row.value[0].ToString();
                                    ChatbotSheetLT[row.index] = row.value[1].ToString();
                                    ChatbotSheetResult[row.index] = row.value[2].ToString();
                                    ChatbotUpdateTime[row.index] = DateTime.Now.ToString("MM-dd HH:mm:dd");
                                }
                            }
                            catch (Exception)
                            {
                                if (ChatbotSheetResult[row.index] == "")
                                {
                                    ChatbotSheetName[row.index] = row.value[0].ToString();
                                    ChatbotSheetLT[row.index] = "0:00:00";
                                    ChatbotSheetResult[row.index] = "0";
                                }
                            }
                        }
                    }
                    Thread.Sleep(60000 * 5);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }
        #endregion

        #endregion
    }
}
