using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace Pink_PinkbeanBot
{
    class LeagueOfLegends
    {
        #region Function
        static string API = "";
        public static string summonerName;

        static string[] MatchDetail = new string[10];
        static string[] GameID = new string[10];
        #endregion

        #region Request
        public static string HttpRequest(string URL)
        {
            WebRequest request = WebRequest.Create(URL + API);
            WebResponse response = request.GetResponse();
            StreamReader stream = new StreamReader(response.GetResponseStream());
            string web = stream.ReadToEnd();

            return web;
        }
        #endregion

        #region ChampionID
        static string ChampionID(int id)
        {
            string ChampionNameKR = "Null";
            string Emoji = "Null";

            switch (id)
            {
                case 266:
                    ChampionNameKR = "아트록스";
                    Emoji = "<:Aatrox:848561151100583976>";
                    break;

                case 412:
                    ChampionNameKR = "쓰레쉬";
                    Emoji = "<:Thresh:848561430692495380>";
                    break;

                case 23:
                    ChampionNameKR = "트린다미어";
                    Emoji = "<:Tryndamere:848561429908946944>";
                    break;

                case 79:
                    ChampionNameKR = "그라가스";
                    Emoji = "<:Gragas:848561149858938921>";
                    break;

                case 69:
                    ChampionNameKR = "카시오페아";
                    Emoji = "<:Cassiopeia:848561150315986984>";
                    break;

                case 136:
                    ChampionNameKR = "아우렐리온 솔";
                    Emoji = "<:AurelionSol:848561150005477376>";
                    break;

                case 13:
                    ChampionNameKR = "라이즈";
                    Emoji = "<:Ryze:848561429325807667>";
                    break;

                case 78:
                    ChampionNameKR = "뽀삐";
                    Emoji = "<:Poppy:848561228439617576>";
                    break;

                case 14:
                    ChampionNameKR = "사이온";
                    Emoji = "<:Sion:848561430353412116>";
                    break;

                case 1:
                    ChampionNameKR = "애니";
                    Emoji = "<:Annie:848561149975724092>";
                    break;

                case 202:
                    ChampionNameKR = "진";
                    Emoji = "<:Jhin:848561150177968208>";
                    break;

                case 43:
                    ChampionNameKR = "카르마";
                    Emoji = "<:Karma:848561227872468993>";
                    break;

                case 111:
                    ChampionNameKR = "노틸러스";
                    Emoji = "<:Nautilus:848561227776393238>";
                    break;

                case 240:
                    ChampionNameKR = "클레드";
                    Emoji = "<:Kled:848561227558551603>";
                    break;

                case 99:
                    ChampionNameKR = "럭스";
                    Emoji = "<:Lux:848561228770967557>";
                    break;

                case 103:
                    ChampionNameKR = "아리";
                    Emoji = "<:Ahri:848561149725245471>";
                    break;

                case 2:
                    ChampionNameKR = "올라프";
                    Emoji = "<:Olaf:848561229105725470>";
                    break;

                case 112:
                    ChampionNameKR = "빅토르";
                    Emoji = "<:Viktor:848561429602500659>";
                    break;

                case 34:
                    ChampionNameKR = "애니비아";
                    Emoji = "<:Anivia:848561150639472640>";
                    break;

                case 27:
                    ChampionNameKR = "신지드";
                    Emoji = "<:Singed:848561430030450699>";
                    break;

                case 86:
                    ChampionNameKR = "가렌";
                    Emoji = "<:Garen:848561150425956392>";
                    break;

                case 127:
                    ChampionNameKR = "리산드라";
                    Emoji = "<:Lissandra:848561228191236098>";
                    break;

                case 57:
                    ChampionNameKR = "마오카이";
                    Emoji = "<:Maokai:848561228238028830>";
                    break;

                case 25:
                    ChampionNameKR = "모르가나";
                    Emoji = "<:Morgana:848561228300550184>";
                    break;

                case 28:
                    ChampionNameKR = "이블린";
                    Emoji = "<:Evelynn:848561150203002911>";
                    break;

                case 105:
                    ChampionNameKR = "피즈";
                    Emoji = "<:Fizz:848561150165254204>";
                    break;

                case 74:
                    ChampionNameKR = "하이머딩거";
                    Emoji = "<:Heimerdinger:848561151088001024>";
                    break;

                case 238:
                    ChampionNameKR = "제드";
                    Emoji = "<:Zed:848560496210214952>";
                    break;

                case 68:
                    ChampionNameKR = "럼블";
                    Emoji = "<:Rumble:848561429917597776>";
                    break;

                case 82:
                    ChampionNameKR = "모데카이저";
                    Emoji = "<:Mordekaiser:848561228065931264>";
                    break;

                case 37:
                    ChampionNameKR = "소나";
                    Emoji = "<:Sona:848561430411345950>";
                    break;

                case 96:
                    ChampionNameKR = "코그모";
                    Emoji = "<:KogMaw:848561227960942602>";
                    break;

                case 55:
                    ChampionNameKR = "카타리나";
                    Emoji = "<:Katarina:848561228732563486>";
                    break;

                case 117:
                    ChampionNameKR = "룰루";
                    Emoji = "<:Lulu:848561228083363900>";
                    break;

                case 22:
                    ChampionNameKR = "애쉬";
                    Emoji = "<:Ashe:848561150395678741>";
                    break;

                case 30:
                    ChampionNameKR = "카서스";
                    Emoji = "<:Karthus:848561227801559071>";
                    break;

                case 12:
                    ChampionNameKR = "알리스타";
                    Emoji = "<:Alistar:848561149992894474>";
                    break;

                case 122:
                    ChampionNameKR = "다리우스";
                    Emoji = "<:Darius:848561150077698048>";
                    break;

                case 67:
                    ChampionNameKR = "베인";
                    Emoji = "<:Vayne:848561430705995806>";
                    break;

                case 110:
                    ChampionNameKR = "바루스";
                    Emoji = "<:Varus:848561430214737930>";
                    break;

                case 77:
                    ChampionNameKR = "우디르";
                    Emoji = "<:Udyr:848561429900689468>";
                    break;

                case 89:
                    ChampionNameKR = "레오나";
                    Emoji = "<:Leona:848561228757598259>";
                    break;

                case 126:
                    ChampionNameKR = "제이스";
                    Emoji = "<:Jayce:848561150387552261>";
                    break;

                case 134:
                    ChampionNameKR = "신드라";
                    Emoji = "<:Syndra:848561429538668595>";
                    break;

                case 80:
                    ChampionNameKR = "판테온";
                    Emoji = "<:Pantheon:848561228908462080>";
                    break;

                case 92:
                    ChampionNameKR = "리븐";
                    Emoji = "<:Riven:848561228355600394>";
                    break;

                case 121:
                    ChampionNameKR = "카직스";
                    Emoji = "<:Khazix:848561227923718165>";
                    break;

                case 42:
                    ChampionNameKR = "코르키";
                    Emoji = "<:Corki:848561150907645992>";
                    break;

                case 268:
                    ChampionNameKR = "아지르";
                    Emoji = "<:Azir:848561149955670037>";
                    break;

                case 51:
                    ChampionNameKR = "케이틀린";
                    Emoji = "<:Caitlyn:848561150147821580>";
                    break;

                case 76:
                    ChampionNameKR = "니달리";
                    Emoji = "<:Nidalee:848561227705745409>";
                    break;

                case 85:
                    ChampionNameKR = "케인";
                    Emoji = "<:Kayn:848561227977326642>";
                    break;

                case 3:
                    ChampionNameKR = "갈리오";
                    Emoji = "<:Galio:848561150798987314>";
                    break;

                case 45:
                    ChampionNameKR = "베이가";
                    Emoji = "<:Veigar:848561429686386709>";
                    break;

                case 432:
                    ChampionNameKR = "바드";
                    Emoji = "<:Bard:848561149884497932>";
                    break;

                case 150:
                    ChampionNameKR = "나르";
                    Emoji = "<:Gnar:848561150428708894>";
                    break;

                case 90:
                    ChampionNameKR = "말자하";
                    Emoji = "<:Malzahar:848561228838076436>";
                    break;

                case 104:
                    ChampionNameKR = "그레이브즈";
                    Emoji = "<:Graves:848561150404198442>";
                    break;

                case 254:
                    ChampionNameKR = "바이";
                    Emoji = "<:Vi:848561430327590924>";
                    break;

                case 10:
                    ChampionNameKR = "케일";
                    Emoji = "<:Kayle:848561227722260481>";
                    break;

                case 39:
                    ChampionNameKR = "이렐리아";
                    Emoji = "<:Irelia:848561150961647616>";
                    break;

                case 64:
                    ChampionNameKR = "리 신";
                    Emoji = "<:LeeSin:848561227756339211>";
                    break;

                case 420:
                    ChampionNameKR = "일라오이";
                    Emoji = "<:Illaoi:848561150374576158>";
                    break;

                case 60:
                    ChampionNameKR = "엘리스";
                    Emoji = "<:Elise:848561150086217728>";
                    break;

                case 106:
                    ChampionNameKR = "볼리베어";
                    Emoji = "<:Volibear:848561430566797323>";
                    break;

                case 20:
                    ChampionNameKR = "누누";
                    Emoji = "<:Nunu:848561228158730250>";
                    break;

                case 4:
                    ChampionNameKR = "트위스티드 페이트";
                    Emoji = "<:TwistedFate:848561430043164702>";
                    break;

                case 24:
                    ChampionNameKR = "잭스";
                    Emoji = "<:Jax:848561150199201833>";
                    break;

                case 102:
                    ChampionNameKR = "쉬바나";
                    Emoji = "<:Shyvana:848561430499426314>";
                    break;

                case 429:
                    ChampionNameKR = "칼리스타";
                    Emoji = "<:Kalista:848561227918737459>";
                    break;

                case 36:
                    ChampionNameKR = "문도 박사";
                    Emoji = "<:DrMundo:848561150115315712>";
                    break;

                case 427:
                    ChampionNameKR = "아이번";
                    Emoji = "<:Ivern:848561150173118494>";
                    break;

                case 131:
                    ChampionNameKR = "다이애나";
                    Emoji = "<:Diana:848561150027366400>";
                    break;

                case 63:
                    ChampionNameKR = "브랜드";
                    Emoji = "<:Brand:848561149980311552>";
                    break;

                case 113:
                    ChampionNameKR = "세주아니";
                    Emoji = "<:Sejuani:848561429598830602>";
                    break;

                case 8:
                    ChampionNameKR = "블라디미르";
                    Emoji = "<:Vladimir:848561430395486239>";
                    break;

                case 154:
                    ChampionNameKR = "자크";
                    Emoji = "<:Zac:848561430407544832>";
                    break;

                case 421:
                    ChampionNameKR = "렉사이";
                    Emoji = "<:RekSai:848561228275646485>";
                    break;

                case 133:
                    ChampionNameKR = "퀸";
                    Emoji = "<:Quinn:848561228187566120>";
                    break;

                case 84:
                    ChampionNameKR = "아칼리";
                    Emoji = "<:Akali:848561150010720326>";
                    break;

                case 163:
                    ChampionNameKR = "탈리야";
                    Emoji = "<:Taliyah:848561430000173066>";
                    break;

                case 18:
                    ChampionNameKR = "트리스타나";
                    Emoji = "<:Tristana:848561430177382430>";
                    break;

                case 120:
                    ChampionNameKR = "헤카림";
                    Emoji = "<:Hecarim:848561151197315102>";
                    break;

                case 15:
                    ChampionNameKR = "시비르";
                    Emoji = "<:Sivir:848561429979725874>";
                    break;

                case 236:
                    ChampionNameKR = "루시안";
                    Emoji = "<:Lucian:848561227990695936>";
                    break;

                case 107:
                    ChampionNameKR = "렝가";
                    Emoji = "<:Rengar:848561228333711410>";
                    break;

                case 19:
                    ChampionNameKR = "워윅";
                    Emoji = "<:Warwick:848561430004498432>";
                    break;

                case 72:
                    ChampionNameKR = "스카너";
                    Emoji = "<:Skarner:848561429921136650>";
                    break;

                case 54:
                    ChampionNameKR = "말파이트";
                    Emoji = "<:Malphite:848561228007735337>";
                    break;

                case 157:
                    ChampionNameKR = "야스오";
                    Emoji = "<:Yasuo:848561430068330536>";
                    break;

                case 101:
                    ChampionNameKR = "제라스";
                    Emoji = "<:Xerath:848561430449356881>";
                    break;

                case 17:
                    ChampionNameKR = "티모";
                    Emoji = "<:Teemo:848561430503751728>";
                    break;

                case 75:
                    ChampionNameKR = "나서스";
                    Emoji = "<:Nasus:848561229169819678>";
                    break;

                case 58:
                    ChampionNameKR = "레넥톤";
                    Emoji = "<:Renekton:848561229001916416>";
                    break;

                case 119:
                    ChampionNameKR = "드레이븐";
                    Emoji = "<:Draven:848561150119116800>";
                    break;

                case 35:
                    ChampionNameKR = "샤코";
                    Emoji = "<:Shaco:848561429476540417>";
                    break;

                case 50:
                    ChampionNameKR = "스웨인";
                    Emoji = "<:Swain:848561429739995186>";
                    break;

                case 91:
                    ChampionNameKR = "탈론";
                    Emoji = "<:Talon:848561430424059904>";
                    break;

                case 40:
                    ChampionNameKR = "잔나";
                    Emoji = "<:Janna:848561150597005352>";
                    break;

                case 115:
                    ChampionNameKR = "직스";
                    Emoji = "<:Ziggs:848560496596353054>";
                    break;

                case 245:
                    ChampionNameKR = "에코";
                    Emoji = "<:Ekko:848561149758931035>";
                    break;

                case 61:
                    ChampionNameKR = "오리아나";
                    Emoji = "<:Orianna:848561227914805299>";
                    break;

                case 114:
                    ChampionNameKR = "피오라";
                    Emoji = "<:Fiora:848562999765893120>";
                    break;

                case 9:
                    ChampionNameKR = "피들스틱";
                    Emoji = "<:FiddleSticks:848561151105433610>";
                    break;

                case 31:
                    ChampionNameKR = "초가스";
                    Emoji = "<:Chogath:848561150291345439>";
                    break;

                case 33:
                    ChampionNameKR = "람머스";
                    Emoji = "<:Rammus:848563840312148018>";
                    break;

                case 7:
                    ChampionNameKR = "르블랑";
                    Emoji = "<:Leblanc:848561227718328321>";
                    break;

                case 16:
                    ChampionNameKR = "소라카";
                    Emoji = "<:Soraka:848561430013935646>";
                    break;

                case 26:
                    ChampionNameKR = "질리언";
                    Emoji = "<:Zilean:848560496260415519>";
                    break;

                case 56:
                    ChampionNameKR = "녹턴";
                    Emoji = "<:Nocturne:848561228158337034>";
                    break;

                case 222:
                    ChampionNameKR = "징크스";
                    Emoji = "<:Jinx:848561150291083264>";
                    break;

                case 83:
                    ChampionNameKR = "요릭";
                    Emoji = "<:Yorick:848561429794521149>";
                    break;

                case 6:
                    ChampionNameKR = "우르곳";
                    Emoji = "<:Urgot:848561430230990858>";
                    break;

                case 203:
                    ChampionNameKR = "킨드레드";
                    Emoji = "<:Kindred:848561228100272169>";
                    break;

                case 21:
                    ChampionNameKR = "미스 포츈";
                    Emoji = "<:MissFortune:848561228875300884>";
                    break;

                case 62:
                    ChampionNameKR = "오공";
                    Emoji = "<:MonkeyKing:848561228104204308>";
                    break;

                case 53:
                    ChampionNameKR = "블리츠크랭크";
                    Emoji = "<:Blitzcrank:848561150492278807>";
                    break;

                case 98:
                    ChampionNameKR = "쉔";
                    Emoji = "<:Shen:848561429421359115>";
                    break;

                case 201:
                    ChampionNameKR = "브라움";
                    Emoji = "<:Braum:848561149963534346>";
                    break;

                case 5:
                    ChampionNameKR = "신 짜오";
                    Emoji = "<:XinZhao:848561430050897920>";
                    break;

                case 29:
                    ChampionNameKR = "트위치";
                    Emoji = "<:Twitch:848561429912223754>";
                    break;

                case 11:
                    ChampionNameKR = "마스터 이";
                    Emoji = "<:MasterYi:848561228770312222>";
                    break;

                case 44:
                    ChampionNameKR = "타릭";
                    Emoji = "<:Taric:848561429547450409>";
                    break;

                case 32:
                    ChampionNameKR = "아무무";
                    Emoji = "<:Amumu:848561150073634846>";
                    break;

                case 41:
                    ChampionNameKR = "갱플랭크";
                    Emoji = "<:Gangplank:848561150781947935>";
                    break;

                case 48:
                    ChampionNameKR = "트런들";
                    Emoji = "<:Trundle:848561429489516615>";
                    break;

                case 38:
                    ChampionNameKR = "카사딘";
                    Emoji = "<:Kassadin:848561228335022090>";
                    break;

                case 161:
                    ChampionNameKR = "벨코즈";
                    Emoji = "<:Velkoz:848561430633381898>";
                    break;

                case 143:
                    ChampionNameKR = "자이라";
                    Emoji = "<:Zyra:848560496248750090>";
                    break;

                case 267:
                    ChampionNameKR = "나미";
                    Emoji = "<:Nami:848561228078645298>";
                    break;

                case 59:
                    ChampionNameKR = "자르반 4세";
                    Emoji = "<:JarvanIV:848561150172856330>";
                    break;

                case 81:
                    ChampionNameKR = "이즈리얼";
                    Emoji = "<:Ezreal:848561150857314344>";
                    break;

                case 350:
                    ChampionNameKR = "유미";
                    Emoji = "<:Yuumi:848561430554214401>";
                    break;

                case 145:
                    ChampionNameKR = "카이사";
                    Emoji = "<:Kaisa:848561150270111784>";
                    break;

                case 518:
                    ChampionNameKR = "니코";
                    Emoji = "<:Neeko:848561228653789196>";
                    break;

                case 142:
                    ChampionNameKR = "조이";
                    Emoji = "<:Zoe:848560496537108564>";
                    break;

                case 498:
                    ChampionNameKR = "자야";
                    Emoji = "<:Xayah:848561429955215360>";
                    break;

                case 517:
                    ChampionNameKR = "사일러스";
                    Emoji = "<:Sylas:848564408498257952>";
                    break;

                case 141:
                    ChampionNameKR = "케인";
                    Emoji = "<:Kayn:848561227977326642>";
                    break;

                case 516:
                    ChampionNameKR = "오른";
                    Emoji = "<:Ornn:848561228405800960>";
                    break;

                case 555:
                    ChampionNameKR = "파이크";
                    Emoji = "<:Pyke:848561228951322686>";
                    break;

                case 164:
                    ChampionNameKR = "카밀";
                    Emoji = "<:Camille:848561149972054026>";
                    break;

                case 246:
                    ChampionNameKR = "키아나";
                    Emoji = "<:Qiyana:848561229211238410>";
                    break;

                case 497:
                    ChampionNameKR = "라칸";
                    Emoji = "<:Rakan:848561228208930816>";
                    break;

                case 777:
                    ChampionNameKR = "요네";
                    Emoji = "<:Yone:848561430328115250>";
                    break;

                case 876:
                    ChampionNameKR = "릴리아";
                    Emoji = "<:Lillia:848561228795478016>";
                    break;

                case 235:
                    ChampionNameKR = "세나";
                    Emoji = "<:Senna:848561430369665054>";
                    break;

                case 875:
                    ChampionNameKR = "세트";
                    Emoji = "<:Sett:848561429295661076>";
                    break;

                case 523:
                    ChampionNameKR = "아펠리오스";
                    Emoji = "<:Aphelios:848561150047551528>";
                    break;

                case 223:
                    ChampionNameKR = "탐 켄치";
                    Emoji = "<:TahmKench:848561429871067167>";
                    break;

                case 360:
                    ChampionNameKR = "사미라";
                    Emoji = "<:Samira:848561429678260244>";
                    break;

                case 887:
                    ChampionNameKR = "그웬";
                    Emoji = "<:Gwen:848561150869504090>";
                    break;

                case 234:
                    ChampionNameKR = "비에고";
                    Emoji = "<:Viego:848561430696296508>";
                    break;

                case 526:
                    ChampionNameKR = "렐";
                    Emoji = "<:Rell:894757973283454976>";
                    break;

                case 711:
                    ChampionNameKR = "벡스";
                    Emoji = "<:Vex:894757526988525648>";
                    break;

                case 221:
                    ChampionNameKR = "제리";
                    Emoji = "<:Zeri:939181704462073916>";
                    break;

                case 166:
                    ChampionNameKR = "아크샨";
                    Emoji = "<:Akshan:960659921572081724>";
                    break;

                default:
                    ChampionNameKR = "Error";
                    Emoji = "Error";
                    break;
            }

            return ChampionNameKR;
        }
        #endregion

        #region Tier
        static string TierEmoji(string tier)
        {
            string Emoji = "";

            switch (tier)
            {
                case "IRON":
                    Emoji = "<:Iron:848827213167460352>";
                    break;

                case "BRONZE":
                    Emoji = "<:Bronze:848827212954337280>";
                    break;

                case "SILVER":
                    Emoji = "<:Silver:848827213029703710>";
                    break;

                case "GOLD":
                    Emoji = "<:Gold:848827212568199209>";
                    break;

                case "PLATINUM":
                    Emoji = "<:Platinum:848827213155663883>";
                    break;

                case "DIAMOND":
                    Emoji = "<:Diamond:848827212647628811>";
                    break;

                case "MASTER":
                    Emoji = "<:Master:848827212672925717>";
                    break;

                case "GRANDMASTER":
                    Emoji = "<:Grandmaster:848827212643303435>";
                    break;

                case "CHALLENGER":
                    Emoji = "<:Challenger:848827213133643818>";
                    break;

                default:
                    Emoji = "UNRANK";
                    break;
            }

            return Emoji;
        }
        #endregion

        #region Rotations
        public static string[] Rotations()
        {
            List<string> name = new List<string>();

            string[] bs1 = LeagueOfLegends.HttpRequest("https://kr.api.riotgames.com/lol/platform/v3/champion-rotations?api_key=").Split(new string[] { "[" }, StringSplitOptions.None);
            string base2 = bs1[1];
            string[] bs2 = base2.Split(new char[] { ',' });
            string base3 = bs2[15];
            string[] bs3 = base3.Split(new char[] { ']' });

            for (int i = 0; i < 15; i++)
            {
                name.Add(ChampionID(Convert.ToInt32(bs2[i])));
            }
            name.Add(ChampionID(Convert.ToInt32(bs3[0])));

            string[] rotations = name.ToArray();

            return rotations;
        }
        #endregion

        #region Summoner
        public static string AccountID()
        {
            string[] bs1 = LeagueOfLegends.HttpRequest("https://kr.api.riotgames.com/lol/summoner/v4/summoners/by-name/" + summonerName + "?api_key=").Split(new string[] { "\"accountId\":" }, StringSplitOptions.None);

            string[] puuid = bs1[1].Split(new string[] { "\"puuid\":" }, StringSplitOptions.None);
            string[] puuid2 = puuid[1].Split(new char[] { '"' });

            string[] name = bs1[1].Split(new string[] { "\"name\":" }, StringSplitOptions.None);
            string[] name2 = name[1].Split(new char[] { '"' });

            summonerName = name2[1];

            return puuid2[1];
        }

        public static string[] SummonerID()
        {
            string[] bs1 = LeagueOfLegends.HttpRequest("https://kr.api.riotgames.com/lol/summoner/v4/summoners/by-name/" + summonerName + "?api_key=").Split(new string[] { "\"id\":" }, StringSplitOptions.None);
            string base2 = bs1[1];
            string[] bs2 = base2.Split(new char[] { '"' });

            return bs2;
        }

        public static string[] SummonerInfo()
        {
            List<string> Info = new List<string>();
            string[] SummonerID = LeagueOfLegends.SummonerID();

            string[] API = LeagueOfLegends.HttpRequest("https://kr.api.riotgames.com/lol/league/v4/entries/by-summoner/" + SummonerID[1] + "?api_key=").Split(new string[] { "RANKED_SOLO_5x5" }, StringSplitOptions.None);

            try
            {
                string[] tier = API[1].Split(new string[] { "\"tier\":" }, StringSplitOptions.None);
                string[] tier2 = tier[1].Split(new char[] { '"' });

                Info.Add(tier2[1]);
                Info.Add(TierEmoji(tier2[1]));

                string[] rank = API[1].Split(new string[] { "\"rank\":" }, StringSplitOptions.None);
                string[] rank2 = rank[1].Split(new char[] { '"' });
                Info.Add(rank2[1]);


                string[] wins = API[1].Split(new string[] { "\"wins\":" }, StringSplitOptions.None);
                string[] wins2 = wins[1].Split(new string[] { "," }, StringSplitOptions.None);
                Info.Add(wins2[0]);

                string[] losses = API[1].Split(new string[] { "\"losses\":" }, StringSplitOptions.None);
                string[] losses2 = losses[1].Split(new string[] { "," }, StringSplitOptions.None);
                Info.Add(losses2[0]);

                string[] leaguePoints = API[1].Split(new string[] { "\"leaguePoints\":" }, StringSplitOptions.None);
                string[] leaguePoints2 = leaguePoints[1].Split(new string[] { "," }, StringSplitOptions.None);
                Info.Add(leaguePoints2[0]);

                string[] info = Info.ToArray();

                return info;
            }
            catch (Exception)
            {
                return null;
            }
        }
        #endregion

        #region Match
        public static void GetGameID()
        {
            List<string> id = new List<string>();

            string accountid = LeagueOfLegends.AccountID();

            string bs1 = LeagueOfLegends.HttpRequest("https://asia.api.riotgames.com/lol/match/v5/matches/by-puuid/" + accountid + "/ids?type=ranked&start=0&count=10&api_key=");

            for (int i = 0; i < 10; i++)
            {
                string[] base1 = bs1.Split(new string[] { "\"," }, StringSplitOptions.None);
                string[] base2 = base1[0 + i].Split(new string[] { "\"" }, StringSplitOptions.None);
                id.Add(base2[1]);
            }

            GameID = id.ToArray();
        }

        public static void MatchAPI()
        {
            GetGameID();

            string FindArray = null;

            List<string> Detail = new List<string>();

            for (int a = 0; a < 10; a++)
            {
                string API = LeagueOfLegends.HttpRequest("https://asia.api.riotgames.com/lol/match/v5/matches/" + GameID[a] + "?api_key=");

                for (int i = 0; i < 10; i++)
                {
                    string[] detailfind = API.Split(new string[] { "{\"assists\":" }, StringSplitOptions.None);
                    int Find = detailfind[1 + i].IndexOf(summonerName);

                    if (Find != -1)
                    {
                        FindArray = detailfind[1 + i];
                    }
                }
                Detail.Add(FindArray);
            }

            MatchDetail = Detail.ToArray();
        }

        #region Champion
        public static string[] Champion()
        {
            MatchAPI();

            List<string> Result = new List<string>();

            for (int i = 0; i < MatchDetail.Length; i++)
            {
                string[] resultline = MatchDetail[i].Split(new string[] { "\"championId\":" }, StringSplitOptions.None);
                string[] resultline2 = resultline[1].Split(new string[] { "," }, StringSplitOptions.None);
                Result.Add(ChampionID(Convert.ToInt32(resultline2[0])));
            }

            string[] result = Result.ToArray();
            return result;
        }
        #endregion

        #region Result
        public static string[] GameResult()
        {
            List<string> Result = new List<string>();

            for (int i = 0; i < 10; i++)
            {
                string[] resultline = MatchDetail[i].Split(new string[] { "\"win\":" }, StringSplitOptions.None);
                string[] resultline2 = resultline[1].Split(new string[] { "}" }, StringSplitOptions.None);

                string KRResult = resultline2[0];

                if (KRResult == "true")
                {
                    KRResult = ":Win:";
                }
                else if (KRResult == "false")
                {
                    KRResult = ":Lose1:";
                }

                Result.Add(KRResult);
            }

            string[] result = Result.ToArray();

            return result;
        }
        #endregion

        #region KDA
        public static string[] Kill()
        {
            List<string> Result = new List<string>();

            for (int i = 0; i < 10; i++)
            {
                string[] resultline = MatchDetail[i].Split(new string[] { "\"kills\":" }, StringSplitOptions.None);
                string[] resultline2 = resultline[1].Split(new string[] { "," }, StringSplitOptions.None);

                Result.Add(resultline2[0]);
            }

            string[] result = Result.ToArray();

            return result;
        }

        public static string[] Death()
        {
            List<string> Result = new List<string>();

            for (int i = 0; i < 10; i++)
            {
                string[] resultline = MatchDetail[i].Split(new string[] { "\"deaths\":" }, StringSplitOptions.None);
                string[] resultline2 = resultline[1].Split(new string[] { "," }, StringSplitOptions.None);

                Result.Add(resultline2[0]);
            }

            string[] result = Result.ToArray();

            return result;
        }

        public static string[] Assist()
        {
            List<string> Result = new List<string>();

            for (int i = 0; i < 10; i++)
            {
                string[] line = MatchDetail[i].Split(new char[] { ',' });

                Result.Add(line[0]);
            }

            string[] result = Result.ToArray();

            return result;
        }
        #endregion

        #endregion
    }
}
