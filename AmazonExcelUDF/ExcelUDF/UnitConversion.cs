using AmazonExcelAddIn.UserLibrary;
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace AmazonExcelUDF.ExcelUDF
{
    public partial class UnitConversion
    {
        #region 长度单位转换器
        [ExcelFunction(Category = "单位换算", IsVolatile = true, IsThreadSafe = true, Description = @"
        长度单位转换器
        公制: 千米(KM)|米(M)|分米(DM)|厘米(CM)|毫米(CM)|微米(UM)|纳米(NM)|皮米(PM)|飞米(FM)|阿米(AM)
        英制: 英寸(IN) | 英尺(FT) | 码(YD) | 英里(MI) | 海里(NMI) | 英寻(EFM) | 弗隆(FUR)
        市制: 里 | 丈 | 尺 | 寸 | 分 | 厘")]
        public static object LengthConverter(
             [ExcelArgument(Description = "输入待转换的值")] double valueConversion,
             [ExcelArgument(Description = "转换前单位")] string unitbeforeConversion,
             [ExcelArgument(Description = "转换后单位")] string unitAfterConversion,
             [ExcelArgument(Description = "是否输出全部转换")] bool isAll = false
            )
        {
            unitbeforeConversion = unitbeforeConversion.ToUpper();
            unitAfterConversion = unitAfterConversion.ToUpper();
            double engFoot = 0.3048;
            Dictionary<string, double> lengthData = new Dictionary<string, double>() {
                {"KM",1000 },{"公里",1000 },
                {"M",1 },{"米",1 },
                {"DM",0.1 },{"分米",0.1 },
                {"CM",0.01 },{"厘米",0.01 },
                {"MM",0.001 },{"毫米",0.001 },
                {"UM",0.000001 },{"微米",0.000001 },
                {"NM",0.000000001 },{"纳米",0.000000001 },
                {"PM",0.000000000001 },{"皮米",0.000000000001 },
                {"FM",0.000000000000001 },{"飞米",0.000000000000001 },
                {"AM",0.000000000000000001 },{"阿米",0.000000000000000001 },
                {"里",500 },{"丈",(double)10 / 3},{"尺",(double)1 / 3 },{"寸",(double)1 / 30 },
                {"分",(double)1 / 300 },{"厘",(double)1 / 3000 },{"毫",(double)1 / 30000 },
                {"FT",engFoot },{"英尺",engFoot },
                {"MI",5280 * engFoot},{"英里",5280 * engFoot},
                {"FUR",660 * engFoot},{"弗隆",660 * engFoot},
                {"YD",3 * engFoot},{"码",3 * engFoot},
                {"IN",engFoot / 12},{"英寸",engFoot / 12},
                {"NMI",1852},{"海里",1852},
                {"EFM",6 * engFoot},{"英寻",6 * engFoot},
            };
            double uval = valueConversion * lengthData[unitbeforeConversion];

            if (isAll)
            {
                return GetAllConversion(lengthData, uval);
            }
            else
            {
                return (uval / lengthData[unitAfterConversion]);
            }


        }
        #endregion

        #region 重量单位转换器
        [ExcelFunction(Category = "单位换算", IsVolatile = true, IsThreadSafe = true, Description = @"
        重量单位转换器
        市制: 担|斤|两|钱
        公制: 吨(T)|千克(KG)|百克(HG)|克拉(CT)|克(G)|厘克(CG)|毫克(MG)|微克(UG)|毫微克(NG)
        金衡制:金衡磅(LBT)|金衡盎司(OZT)|英钱(DWT)|金衡格令(GRT)
        常衡制:长吨(LT)|短吨(ST)|英担(BRICWT)|美担(USCWT)|英石(BRIST)|磅(LB)|盎司(OZ)|打兰(DR)|格令(GR)")]
        public static object WeightConverter(
             [ExcelArgument(Description = "输入待转换的值")] double valueConversion,
             [ExcelArgument(Description = "转换前单位")] string unitbeforeConversion,
             [ExcelArgument(Description = "转换后单位")] string unitAfterConversion,
             [ExcelArgument(Description = "是否输出全部转换")] bool isAll = false)
        {
            unitbeforeConversion = unitbeforeConversion.ToUpper();
            unitAfterConversion = unitAfterConversion.ToUpper();
            double avdpPound = 0.45359237;
            Dictionary<string, double> weightData = new Dictionary<string, double>() {
                {"T", 1000 },{"吨",1000 },
                { "KG", 1 },{ "千克", 1 },
                { "HG", 0.01 },{ "百克", 0.01 },
                { "CT", 0.005 },{ "克拉", 0.001 },
                { "G", 0.001 },{ "克", 0.001 },
                { "CG", 0.00001 },{ "厘克", 0.00001 },
                { "MG", 0.000001 },{ "毫克", 0.000001 },
                { "UG",0.000000001 },{ "微克", 0.000000001 },
                { "NG", 0.000000000001 },{ "毫微克", 0.000000000001 },
                { "担", 50 },{ "斤", 0.5 },{ "两", 0.05 },{ "钱", 0.005 },
                { "LB", avdpPound },{ "磅", avdpPound },
                { "LT", 2240 * avdpPound },{ "长吨", 2240 * avdpPound },
                { "ST", 2000 * avdpPound },{ "短吨", 2000 * avdpPound },
                { "BRICWT", 112 * avdpPound },{ "英担", 112 * avdpPound },
                { "USCWT", 100 * avdpPound },{ "美担", 100 * avdpPound },
                { "BRIST", 14 * avdpPound },{ "英石", 14 * avdpPound },
                { "OZ", avdpPound / 16 },{ "盎司", avdpPound / 16},
                { "DR", avdpPound / 256 },{ "打兰", avdpPound / 256 },
                { "GR", avdpPound / 7000 },{ "格令", avdpPound / 7000 },
                { "LBT", 5760 * avdpPound },{ "金衡磅", 5760 * avdpPound },
                { "OZT", 480 * avdpPound },{ "金衡盎司", 480 * avdpPound },
                { "DWT", 24 * avdpPound },{ "英钱", 24 * avdpPound },
                { "GRT", 24 * avdpPound },{ "金衡格令", avdpPound / 7000 }
            };
            double uval = valueConversion * weightData[unitbeforeConversion];
            if (isAll)
            {
                return GetAllConversion(weightData, uval);
            }
            else
            {
                return (uval / weightData[unitAfterConversion]);
            }
        }
        #endregion

        #region 面积单位转换器
        [ExcelFunction(Category = "单位换算", IsVolatile = true, IsThreadSafe = true, Description = @"
        面积单位转换器
        市制: 顷|亩|分|平方尺|平方寸
        公制: 公顷(HA)|公亩(ARE)|平方公里(KM2)|平方米(M2)|平方分米(DM2)|平方厘米(CM3)|平方毫米(MM2)
        英制: 英亩(ACRE)|平方英里(SQMI)|平方码(SQYD)|平方英尺(SQFT)|平方英寸(SQIN)|平方竿(SQRD)")]
        public static object AreaConverter(
             [ExcelArgument(Description = "输入待转换的值")] double valueConversion,
             [ExcelArgument(Description = "转换前单位")] string unitbeforeConversion,
             [ExcelArgument(Description = "转换后单位")] string unitAfterConversion,
             [ExcelArgument(Description = "是否输出全部转换")] bool isAll = false)
        {
            unitbeforeConversion = unitbeforeConversion.ToUpper();
            unitAfterConversion = unitAfterConversion.ToUpper();
            double engSquareFoot = (0.3048 * 0.3048);
            double usSquareRod = (16.5 * 16.5 * engSquareFoot);
            Dictionary<string, double> areaData = new Dictionary<string, double>() {
                { "顷", ((double)10000 / 15) * 100 },{ "亩", ((double)10000 / 15) * 1 },{ "分", ((double)10000 / 15) * 0.1 },
                { "平方尺", ((double)10000 / 9) * 0.0001 },{ "平方寸", ((double)10000 / 9) * 0.000001 },
                { "HA", (100 * 100) },{ "公顷", (100 * 100) },
                { "ARE", (10 * 10) },{ "公亩",  (10 * 10) },
                { "KM2", (1000 * 1000) },{ "平方公里",(1000 * 1000) },
                { "M2", 1 },{ "平方米", 1 },
                { "DM2", (0.1 * 0.1) },{ "平方分米", (0.1 * 0.1) },
                { "CM2", (0.01 * 0.01) },{ "平方厘米", (0.01 * 0.01) },
                { "MM2", (0.001 * 0.001) },{ "平方毫米", (0.001 * 0.001) },
                { "SQFT", engSquareFoot },{ "平方英尺", engSquareFoot },
                { "SQYD", (3 * 3 * engSquareFoot) },{ "平方码", (3 * 3 * engSquareFoot) },
                { "SQRD", usSquareRod },{ "平方竿", usSquareRod },
                { "ACRE",160 * usSquareRod },{ "英亩", 160 * usSquareRod },
                { "SQMI", (5280 * 5280 * engSquareFoot) },{ "平方英里", (5280 * 5280 * engSquareFoot) },
                { "SQIN", (engSquareFoot / (12 * 12)) },{ "平方英寸", (engSquareFoot / (12 * 12)) }
            };
            double uval = valueConversion * areaData[unitbeforeConversion];
            if (isAll)
            {
                return GetAllConversion(areaData, uval);
            }
            else
            {
                return (uval / areaData[unitAfterConversion]);
            }
        }
        #endregion

        #region 体积单位转换器
        [ExcelFunction(Category = "单位换算", IsVolatile = true, IsThreadSafe = true, Description = @"
        体积单位转换器,更多帮助请设置 isHelp = True; 
        公制尺寸: 立方米(M3)|公石(HL)|十升(DAL)|立方分米(DM3)|立方厘米(CM3)|立方毫米(MM3)
        英制尺寸: 立方英尺(CUFT)|立方英寸(CUIN)|立方码(CUYD)|亩英尺(ACFT)
        公制液量: 升(L)|分升(DL)|毫升(ML)|厘升(CL)|微升(UL)
        ")]
        public static object VolumeConverter(
             [ExcelArgument(Description = "输入待转换的值")] double valueConversion,
             [ExcelArgument(Description = "转换前单位")] string unitbeforeConversion,
             [ExcelArgument(Description = "转换后单位")] string unitAfterConversion,
             [ExcelArgument(Description = "是否输出全部转换")] bool isAll = false,
             [ExcelArgument(Description = "是否显示帮助")] bool isHelp = false)
        {
            if (isHelp)
            {
                return new object[,]
                {
                    { "公制尺寸: 立方米(M3)|公石(HL)|十升(DAL)|立方分米(DM3)|立方厘米(CM3)|立方毫米(MM3)" },
                    { "英制尺寸: 立方英尺(CUFT)|立方英寸(CUIN)|立方码(CUYD)|亩英尺(ACFT)" },
                    { "公制液量: 升(L)|分升(DL)|毫升(ML)|厘升(CL)|微升(UL)" },
                    { "美制液量: 桶[美液](USLBAR)|加仑[美液](USLGAL)|夸脱[美液](USLQT)|品脱[美液](USLPT)|及耳[美液](USLGI)|液量盎司[美液](USLFLOZ)|液量打兰[美液](USLFLDR)|量滴[美液](USLMIN)" },
                    { "美制干量: 桶[美干](USDBAR)|蒲式耳[美干](USDBU)|配克[美干](USDPK)|夸脱[美干](USDQT)|品脱[美干](USDPT)" },
                    { "英制液量和干量: 桶[英](BRIBAR)|蒲式耳[英](BRIBU)|加仑[英](BRIGAL)|夸脱[英](BRIQT)|液量盎司[英](BRIFLOZ)" },
                    { "美制烹调制式: 汤勺[美](USCTBS)|调羹[美](USCTSP)|杯[美](USCFLOZ)" },
                    { "公制烹调制式: 汤勺[公](MCTBS)|调羹[公](MCTSP)" },
                };
            }
            unitbeforeConversion = unitbeforeConversion.ToUpper();
            unitAfterConversion = unitAfterConversion.ToUpper();
            double uscCubicInch = 0.016387064;
            double uslGallon = 231 * uscCubicInch;
            double usdBushel = 2150.42 * uscCubicInch;
            double uslFluidOunce = uslGallon / 128;
            double briGallon = 4.54609;
            Dictionary<string, double> volumeData = new Dictionary<string, double>() {
                { "M3", 1000 },{ "立方米", 1000 },
                { "HL", 100 },{ "公石",  100 },
                { "DAL", 10 },{ "十升",10 },
                { "L", 1 },{ "升", 1 },
                { "DM3", 1 },{ "立方分米", 1 },
                { "DL", 0.1 },{ "分升", 0.1 },
                { "CL", 0.01 },{ "厘升", 0.01 },
                { "ML", 0.001 },{ "毫升", 0.001 },
                { "CM3", 0.001 },{ "立方厘米", 0.001 },
                { "UL", 0.000001 },{ "微升", 0.000001 },
                { "MM3", 0.000001 },{ "立方毫米", 0.000001 },
                { "CUIN", uscCubicInch },{ "立方英寸", uscCubicInch },
                { "CUFT", 1728 * uscCubicInch },{ "立方英尺", 1728 * uscCubicInch },
                { "ACFT", 43560 * 1728 * uscCubicInch },{ "亩英尺", 43560 * 1728 * uscCubicInch },
                { "CUYD", 27 * 1728 * uscCubicInch },{ "立方码", 27 * 1728 * uscCubicInch },
                { "USLGAL", uslGallon },{ "加仑[美液]", uslGallon },
                { "USLBAR", 42 * uslGallon },{ "桶[美液]", 42 * uslGallon },
                { "USLQT", uslGallon / 4  },{ "夸脱[美液]", uslGallon / 4 },
                { "USLPT", uslGallon / 8  },{ "品脱[美液]", uslGallon / 8 },
                { "USLGI", uslGallon / 32  },{ "及耳[美液]", uslGallon / 32 },
                { "USLFLOZ", uslFluidOunce  },{ "液量盎司[美液]", uslFluidOunce },
                { "USLFLDR", uslGallon / 1024  },{ "液量打兰[美液]", uslGallon / 1024 },
                { "USLMIN", uslGallon / 61440  },{ "量滴[美液]", uslGallon / 61440 },
                { "USDBAR", 7056 * uscCubicInch  },{ "桶[美干]", 7056 * uscCubicInch },
                { "USDBU", usdBushel  },{ "蒲式耳[美干]", usdBushel },
                { "USDPK", usdBushel / 4 },{ "配克[美干]", usdBushel / 4 },
                { "USDQT", usdBushel / 32 },{ "夸脱[美干]", usdBushel / 32 },
                { "USDPT", usdBushel / 64 },{ "品脱[美干]", usdBushel / 64 },
                { "MCTBS", 0.015 },{ "汤勺[公]", 0.015 },
                { "MCTSP", 0.005 },{ "调羹[公]", 0.005 },
                { "USCTBS", uslFluidOunce / 2 },{ "汤勺[美]", uslFluidOunce / 2 },
                { "USCTSP", uslFluidOunce / 6 },{ "调羹[美]", uslFluidOunce / 6 },
                { "USCFLOZ", 8 * uslFluidOunce },{ "杯[美]", 8 * uslFluidOunce },
                { "BRIGAL", briGallon },{ "加仑[英]", briGallon },
                { "BRIBAR", 36 * briGallon },{ "桶[英]", 36 * briGallon },
                { "BRIBU", 8 * briGallon },{ "蒲式耳[英]", 8 * briGallon },
                { "BRIQT",  briGallon / 8},{ "夸脱[英]", briGallon / 8 },
                { "BRIFLOZ",  briGallon / 160},{ "液量盎司[英]", briGallon / 160 },
            };
            double uval = valueConversion * volumeData[unitbeforeConversion];
            if (isAll)
            {
                return GetAllConversion(volumeData, uval);
            }
            else
            {
                return (uval / volumeData[unitAfterConversion]);
            }
        }
        #endregion

        #region 压力单位转换器
        [ExcelFunction(Category = "单位换算", IsVolatile = true, IsThreadSafe = true, Description = @"
        压力单位转换器
        巴(BAR)|千帕(KPA)|百帕(HPA)|毫巴(MBAR)|帕斯卡{N/M2}(PA)|标准大气压(ATM)
        磅力/平方英尺{lbf/ft2}(PSF)|磅力/平方英寸{lbf/in2}(PSI)|公斤力/平方厘米{kgf/cm2}(KSC)|公斤力/平方米{kgf/m2}(KSM)
        英吋汞柱(INHG)|毫米汞柱(MMHG)|毫米水柱(MMH2O)")]
        public static object PressurehConverter(
             [ExcelArgument(Description = "输入待转换的值")] double valueConversion,
             [ExcelArgument(Description = "转换前单位")] string unitbeforeConversion,
             [ExcelArgument(Description = "转换后单位")] string unitAfterConversion,
             [ExcelArgument(Description = "是否输出全部转换")] bool isAll = false
            )
        {

            unitbeforeConversion = unitbeforeConversion.ToUpper();
            unitAfterConversion = unitAfterConversion.ToUpper();
            Dictionary<string, double> pressurehData = new Dictionary<string, double>() {
                {"BAR",100000 },{"巴",100000 },
                {"KPA",1000 },{"千帕",1000 },
                {"HPA",100 },{"百帕",100 },
                {"MBAR",100 },{"毫巴",100 },
                {"PA",1 },{"帕斯卡",1 },
                {"ATM",101325 },{"标准大气压",101325 },
                {"MMH2O",(double)1 / 0.101972 },{"毫米水柱",(double)1 / 0.101972 },
                {"MMHG",(double)101325 / 760 },{"毫米汞柱",(double)101325 / 760 },
                {"INHG",25.4 * ((double)101325 / 760) },{"英吋汞柱",25.4 * ((double)101325 / 760) },
                {"PSI",6894.757 },{"磅力/平方英寸",6894.757 },
                {"PSF",6894.757 / 144 },{"磅力/平方英尺",6894.757 / 144 },
                {"KSC",98066.5 },{"公斤力/平方厘米",98066.5 },
                {"KSM",9.80665 },{"公斤力/平方米",9.80665 },
            };
            double uval = valueConversion * pressurehData[unitbeforeConversion];
            if (isAll)
            {
                return GetAllConversion(pressurehData, uval);
            }
            else
            {
                return (uval / pressurehData[unitAfterConversion]);
            }
        }
        #endregion

        #region 温度单位转换器
        [ExcelFunction(Category = "单位换算", IsVolatile = true, IsThreadSafe = true, Description = @"
        温度单位转换器
        摄氏度(C)|华氏度(F)|开氏度(K)|兰氏度(Ra)|列氏度(Re)

　　    精确的测量表明：零摄氏度(冰点)比水的三相点低0.01度；
　　    K＝5/9（°F+459.67） K=℃+273.15
　　    n℃=(5/9·n+32) °F n°F=[(n-32)×5/9]℃
　　    1°F=5/9℃（温度差）")]
        public static object TemperatureConversion(
            [ExcelArgument(Description = "输入待转换的值")] double valueConversion,
            [ExcelArgument(Description = "转换前单位")] string unitbeforeConversion,
            [ExcelArgument(Description = "转换后单位")] string unitAfterConversion,
            [ExcelArgument(Description = "是否输出全部转换")] bool isAll = false
            )
        {
            unitbeforeConversion = unitbeforeConversion.ToUpper();
            unitAfterConversion = unitAfterConversion.ToUpper();
            double tempC = 0, tempK = 0, tempF = 0, tempRa = 0, tempRe = 0;
            switch (unitbeforeConversion)
            {
                case "C":
                case "摄氏度":
                    if (!(valueConversion < -273.15))
                    {
                        tempC = valueConversion;
                        tempK = tempC + 273.15;
                        tempF = 32 + (tempC * 9 / 5);
                        tempRa = tempK * 1.8;
                        tempRe = tempC / 1.25;
                    }
                    break;
                case "F":
                case "华氏度":
                    if (!(valueConversion < -459.666666))
                    {
                        tempF = valueConversion;
                        tempC = (tempF - 32) * 5 / 9;
                        tempK = tempC + 273.15;
                        tempRa = tempK * 1.8;
                        tempRe = tempC / 1.25;
                    }
                    break;
                case "K":
                case "开氏度":
                    if (!(valueConversion < 0))
                    {
                        tempK = valueConversion;
                        tempC = tempK - 273.15;
                        tempF = 32 + (tempC * 9 / 5);
                        tempRa = tempK * 1.8;
                        tempRe = tempC / 1.25;
                    }
                    break;
                case "Ra":
                case "兰氏度":
                    if (!(valueConversion < 0))
                    {
                        tempRa = valueConversion;
                        tempK = tempRa / 1.8;
                        tempC = tempK - 273.15;
                        tempF = 32 + (tempC * 9 / 5);
                        tempRe = tempC / 1.25;
                    }
                    break;
                case "Re":
                case "列氏度":
                    if (!(valueConversion < -218.5199999999))
                    {
                        tempRe = valueConversion;
                        tempC = tempRe * 1.25;
                        tempK = tempC + 273.15;
                        tempF = 32 + (tempC * 9 / 5);
                        tempRa = tempK * 1.8;
                    }
                    break;
            }
            Dictionary<string, double> temperatureData = new Dictionary<string, double>() {
                  {"C",tempC },
                  {"摄氏度",tempC },
                  {"F",tempF },
                  {"华氏度",tempF },
                  {"K",tempK },
                  {"开氏度",tempK },
                  {"Ra",tempRa },
                  {"兰氏度",tempRa },
                  {"Re",tempRe },
                  {"列氏度",tempRe }
            };
            if (isAll)
            {
                return GetAllConversion(temperatureData);
            }
            else
            {
                return temperatureData[unitAfterConversion];
            }
        }
        #endregion

        #region 功率单位转换器
        [ExcelFunction(Category = "单位换算", IsVolatile = true, IsThreadSafe = true, Description = @"
        功率单位转换器
        瓦(W)|千瓦(KW)|英制马力(HP)|米制马力(PS)
        公斤·米/秒(kg·m/s)|千卡/秒(kcal/s)
        英热单位/秒(Btu/s)|英尺·磅/秒(ft·lb/s)
        焦耳/秒(J/s)|牛顿·米/秒(N·m/s)")]
        public static object PowerConverter(
             [ExcelArgument(Description = "输入待转换的值")] double valueConversion,
             [ExcelArgument(Description = "转换前单位")] string unitbeforeConversion,
             [ExcelArgument(Description = "转换后单位")] string unitAfterConversion,
             [ExcelArgument(Description = "是否输出全部转换")] bool isAll = false
            )
        {
            unitbeforeConversion = unitbeforeConversion.ToUpper();
            unitAfterConversion = unitAfterConversion.ToUpper();
            Dictionary<string, double> powerData = new Dictionary<string, double>() {
                {"W", 0.001 },{"瓦", 0.001 },
                {"KW",1 },{"千瓦",1 },
                {"HP", 0.745712172 },{"英制马力", 0.745712172 },
                {"PS", 0.7352941 },{"米制马力", 0.7352941 },
                {"kg·m/s", 0.0098039215 },{"公斤·米/秒", 0.0098039215 },
                {"kcal/s", 4.1841004 },{"千卡/秒", 4.1841004 },
                {"Btu/s", 1.05507491 },{"英热单位/秒", 1.05507491},
                {"ft·lb/s", 0.0013557483731 },{"英尺·磅/秒", 0.0013557483731 },
                {"J/s", 0.001 },{"焦耳/秒", 0.001 },
                {"N·m/s", 0.001 },{"牛顿·米/秒", 0.001 },
            };
            double uval = valueConversion * powerData[unitbeforeConversion];
            if (isAll)
            {
                return GetAllConversion(powerData, uval);
            }
            else
            {
                return (uval / powerData[unitAfterConversion]);
            }
        }
        #endregion

        #region 速度单位转换器
        [ExcelFunction(Category = "单位换算", IsVolatile = true, IsThreadSafe = true, Description = @"
        速度单位转换器
        节(KN)|英里/小时(MP/H)|公里/小时(KM/H)

        科学上用速度来表示物体运动的快慢")]
        public static object SpeedConversion(
            [ExcelArgument(Description = "输入待转换的值")] double valueConversion,
            [ExcelArgument(Description = "转换前单位")] string unitbeforeConversion,
            [ExcelArgument(Description = "转换后单位")] string unitAfterConversion,
            [ExcelArgument(Description = "是否输出全部转换")] bool isAll = false
            )
        {
            unitbeforeConversion = unitbeforeConversion.ToUpper();
            unitAfterConversion = unitAfterConversion.ToUpper();
            double tempKN = 0, tempMPH = 0, tempKMH = 0;
            switch (unitAfterConversion)
            {
                case "KN":
                case "节":
                    tempKN = valueConversion;
                    tempMPH = (tempKN * 1.61) / 1.85;
                    tempKMH = tempKN / 1.85;
                    break;
                case "MP/H":
                case "英里/小时":
                    tempMPH = valueConversion;
                    tempKN = tempMPH * 1.85 * 0.621;
                    tempKMH = tempMPH * 0.621;
                    break;
                case "KM/H":
                case "公里/小时":
                    tempKMH = valueConversion;
                    tempKN = tempKMH * 1.85;
                    tempMPH = tempKMH * 1.61;
                    break;
            }
            Dictionary<string, double> speedData = new Dictionary<string, double>() {
                  {"KN",tempKN },
                  {"节",tempKN },
                  {"MP/H",tempMPH },
                  {"英里/小时",tempMPH },
                  {"KM/H",tempKMH },
                  {"公里/小时",tempKMH }
            };
            if (isAll)
            {
                return GetAllConversion(speedData);
            }
            else
            {
                return speedData[unitbeforeConversion];
            }
        }
        #endregion

        #region 进制单位转换器
        [ExcelFunction(Category = "单位换算", IsVolatile = true, IsThreadSafe = true, Description = @"进制单位转换器，62进制内随意转换")]
        public static object AnyConversion(
            [ExcelArgument(Description = "输入待转换的值")] string valueConversion,
            [ExcelArgument(Description = "转换前进制")] int unitbeforeConversion,
            [ExcelArgument(Description = "转换后进制")] int unitAfterConversion
            )
        {
            valueConversion = valueConversion.ToLower();
            if (unitbeforeConversion == 10)
            {
                return ConvertIntToAny(Convert.ToInt32(valueConversion), unitAfterConversion);
            }
            else
            {
                long tempInt = ConvertAnyToInt(valueConversion, unitbeforeConversion);
                return ConvertIntToAny(Convert.ToInt32(tempInt), unitAfterConversion);
            }
        }
        /// <summary>
        /// (62进制内)10进制转换为指定的进制形式字符串
        /// </summary>
        /// <param name="number">待转换的数字</param>
        /// <param name="coverindex">需要转换的进制（必须在62以内）</param>
        /// <returns>返回 转换进制后的字符串</returns>
        public static string ConvertIntToAny(long number, int coverindex)
        {
            //进制索引表
            string mapcode = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
            //
            string cutmap = mapcode.Substring(0, coverindex);
            //
            List<string> result = new List<string>();
            long t = number;
            int length = cutmap.Length;
            while (t > 0)
            {
                var mod = t % length;
                t = Math.Abs(t / length);
                var character = cutmap[System.Convert.ToInt32(mod)].ToString();
                result.Insert(0, character);
            }
            return string.Join("", result.ToArray());
        }
        /// <summary>
        /// (62进制内)指定的进制形式字符串转换为10进制数字
        /// </summary>
        /// <param name="str">待转换的字符串</param>
        /// <param name="coverindex">字符串对应的进制（必须在62以内）</param>
        /// <returns>返回 十进制数字</returns>
        public static long ConvertAnyToInt(string str, int coverindex)
        {
            string mapcode = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string cutmap = mapcode.Substring(0, coverindex);
            long result = 0;
            int j = 0;
            int length = cutmap.Length;
            foreach (var ch in new string(str.ToCharArray().Reverse().ToArray()))
            {
                if (cutmap.Contains(ch))
                {
                    result += cutmap.IndexOf(ch) * ((long)Math.Pow(length, j));
                    j++;
                }
            }
            return result;
        }
        #endregion

        #region RGB转换器
        [ExcelFunction(Category = "单位换算", Description = "不同颜色间表示方法间的转换")]
        public static object RGBConversion(
            [ExcelArgument(Description = "输入R值，范围0-255")] object inputR,
            [ExcelArgument(Description = "输入G值，范围0-255")] object inputG,
            [ExcelArgument(Description = "输入B值，范围0-255")] object inputB,
            [ExcelArgument(Description = "输出  0: HTML[默认]  1:OLE")] int output = 0

        )
        {
            try
            {
                int R = Convert.ToInt32(inputR.ToString().Trim());
                int G = Convert.ToInt32(inputG.ToString().Trim());
                int B = Convert.ToInt32(inputB.ToString().Trim());

                switch (output)
                {
                    case 0:
                        return ColorTranslator.ToHtml(Color.FromArgb(255, R, G, B));
                    case 1:
                        return ColorTranslator.ToOle(Color.FromArgb(255, R, G, B));
                    default:
                        return ExcelError.ExcelErrorNA;
                }
            }
            catch (Exception)
            {
                return ExcelError.ExcelErrorNA;
            }
        }
        #endregion

        #region Color 转换器

        [ExcelFunction(Category = "单位换算", Description = "不同颜色间表示方法间的转换")]
        public static object ColorConversion(
            [ExcelArgument(Description = "输入 Color 属性[HTML || OLE]")] object input,
            [ExcelArgument(Description = "输出  0: HTML[默认]  1:OLE  2:RGB")] int output = 0
        )
        {
            Color color;
            try
            {
                if (input.ToString().IndexOf("#") > -1)
                {
                    color = ColorTranslator.FromHtml(input.ToString());
                }
                else
                {
                    color = ColorTranslator.FromOle(Convert.ToInt32(input.ToString().Trim()));
                }

                switch (output)
                {
                    case 0:
                        return ColorTranslator.ToHtml(Color.FromArgb(255, color.R, color.G, color.B));
                    case 1:
                        return ColorTranslator.ToOle(color);
                    case 2:
                        return $"{color.R},{color.G},{color.B}";
                    default:
                        return ExcelError.ExcelErrorNA;
                }
            }
            catch (Exception)
            {
                return ExcelError.ExcelErrorNA;
            }
        }
        # endregion

        #region ASCII 转换器
        [ExcelFunction(Category = "单位换算", Description = "字符转ASCII编号")]
        public static object AsciiConversion(
             [ExcelArgument(Description = "输入所要查找的单个字符或者ASCII")] object input,
             [ExcelArgument(Description = "输入的类型  0: ASCII[默认]  1:Char")] int type = 0
        )
        {
            switch (type)
            {
                case 0:
                    int asciiCode = Convert.ToInt32(input.ToString().Trim());
                    if (asciiCode >= 0 && asciiCode <= 255)
                    {
                        ASCIIEncoding asciiEncoding = new ASCIIEncoding();
                        byte[] byteArray = new byte[] { (byte)asciiCode };
                        string strCharacter = asciiEncoding.GetString(byteArray);
                        return (strCharacter);
                    }
                    else
                    {
                        return ExcelError.ExcelErrorNA;
                    }

                case 1:
                    if (input.ToString().Length == 1)
                    {
                        ASCIIEncoding asciiEncoding = new ASCIIEncoding();
                        int intAsciiCode = (int)asciiEncoding.GetBytes(input.ToString())[0];
                        return (intAsciiCode);
                    }
                    else
                    {
                        return ExcelError.ExcelErrorNA;
                    }
                default:
                    return ExcelError.ExcelErrorNA;
            }
        }
        #endregion

        #region 普通日期转UnixTimeStamp
        [ExcelFunction(Category = "单位换算", Description = "普通日期转UnixTimeStamp")]
        public static object DateTimeToUnixTimeStamp(
            [ExcelArgument(Description = "输入UnixTimeStamp")] DateTime inputDateTime,
            [ExcelArgument(Description = "是否精确到秒，TRUE为秒，FALSE为毫秒")] bool isSecond
        )
        {
            DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new DateTime(1970, 1, 1, 0, 0, 0, 0));
            long unixTime = (inputDateTime.Ticks - startTime.Ticks) / 10000;
            Common.ChangeNumberFormat("0");
            return isSecond ? unixTime / 1000 : unixTime;
        }
        #endregion

        #region UnixTimeStamp转普通日期
        [ExcelFunction(Category = "单位换算", Description = "UnixTimeStamp转普通日期")]
        public static object UnixTimeStampToDateTime(
           [ExcelArgument(Description = "输入UnixTimestamp")] Int64 inputUnixTimeStamp,
           [ExcelArgument(Description = "日期格式")] string Format = "yyyy-mm-dd hh:mm:ss"
        )
        {
            if (inputUnixTimeStamp.ToString().Length == 10)
            {
                inputUnixTimeStamp *= 1000;
            }
            DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new DateTime(1970, 1, 1));
            DateTime time = startTime.AddMilliseconds(inputUnixTimeStamp);
            Common.ChangeNumberFormat(Format);
            return time;
        }
        #endregion

        public static string[,] GetAllConversion(Dictionary<string, double> Data, double uval = 0)
        {
            Regex regChina = new Regex("^[^\x00-\xFF]");
            Regex regEnglish = new Regex("^[a-zA-Z]");

            int englishIndex = 0;
            int chinaIndex = 0;
            foreach (string key in Data.Keys)
            {
                if (regEnglish.IsMatch(key))
                {
                    englishIndex += 1;
                }
                else if (regChina.IsMatch(key))
                {
                    chinaIndex += 1;
                }
            }

            int count;
            if (englishIndex > chinaIndex)
            {
                count = englishIndex;
            }
            else
            {
                count = chinaIndex;
            }

            string[,] allConversion = new string[count, 4];
            englishIndex = 0;
            chinaIndex = 0;
            foreach (string key in Data.Keys)
            {
                LoggerHelper.Debug($"{key} -> {uval / Data[key]}");
                if (regEnglish.IsMatch(key))
                {
                    allConversion[englishIndex, 2] = key;
                    if (uval == 0)
                    {
                        allConversion[englishIndex, 3] = Data[key].ToString();
                    }
                    else
                    {
                        allConversion[englishIndex, 3] = (uval / Data[key]).ToString();
                    }
                    englishIndex += 1;
                }
                else if (regChina.IsMatch(key))
                {
                    allConversion[chinaIndex, 0] = key;
                    if (uval == 0)
                    {
                        allConversion[chinaIndex, 1] = Data[key].ToString();
                    }
                    else
                    {
                        allConversion[chinaIndex, 1] = (uval / Data[key]).ToString();
                    }
                    chinaIndex += 1;
                }
            }
            return allConversion;
        }
    }
}