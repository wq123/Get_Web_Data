using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack;

using System.Data;
using System.Reflection;

using Spire.Xls;

namespace 爬数据
{
    class Program
    {

        static void Main(string[] args)
        {
            string url = "";
            List<HouseInfo> tmplist = new List<HouseInfo>();

            int pageNum = 0;
            HtmlNodeCollection AccountToken = null;
            int index = 0;
            // 网址+页码循环抓取数据，直到数据抓完
            do
            {
                pageNum++;
                url = @"http://newhouse.bb.house365.com/house/p-" + pageNum;

                // 使用HtmlAgilityPack解析页面html
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(GetPageSource(url, "utf-8"));

                string xPathStr = "//div[@class='mc']/div[@class='mc_list clearfix']";
                if (pageNum > 1) AccountToken.Clear();//先清空
                AccountToken = doc.DocumentNode.SelectNodes(xPathStr);    //所有找到的节点都是一个集合

                if (AccountToken != null && AccountToken.Count > 1)
                {
                    for (int i = 0; i < AccountToken.Count; i++)
                    {
                        try
                        {
                            HouseInfo tmpmodel = new HouseInfo();
                            // 基本信息
                            tmpmodel.Community = AccountToken[i].SelectNodes("//div[@class='tit']/h3/a")[0].InnerText.Trim();
                            tmpmodel.Area = AccountToken[i].SelectNodes("//div[@class='tit']/span[1]/a")[0].InnerText.Trim();
                            tmpmodel.Position = AccountToken[i].SelectNodes("//div[@class='yh_info f_s']/p[2]")[0].InnerHtml.Split('<')[0];
                            tmpmodel.Price = AccountToken[i].SelectNodes("//div[@class='xiang_price']/span[@class='orange f20 shengluehao']")[0].InnerText;
                            //tmpmodel.Measure = "未知";
                            if (AccountToken[i].SelectNodes("//div[@class='yh_info f_s']/p/span[@class='text_underline']/a") != null)
                                tmpmodel.Measure = AccountToken[i].SelectNodes("//div[@class='yh_info f_s']/p/span[@class='text_underline']/a")[0].InnerText;
                            else
                                tmpmodel.Measure = "未知";
                            // 开盘时间
                            string urlInfo = AccountToken[i].SelectNodes("//div[@class='tit']/h3/a")[0].Attributes["href"].Value;
                            HtmlDocument docInfo = new HtmlDocument();
                            docInfo.LoadHtml(GetPageSource(urlInfo, "utf-8"));
                            if (docInfo.DocumentNode.SelectNodes("//div[@class='w510 fr']/div[4]/span") == null)
                                tmpmodel.OpenTime = "待定";
                            else
                                tmpmodel.OpenTime = docInfo.DocumentNode.SelectNodes("//div[@class='w510 fr']/div[4]/span")[0].InnerText;

                            // 详细信息
                            urlInfo += "intro/";
                            docInfo = new HtmlDocument();
                            docInfo.LoadHtml(GetPageSource(urlInfo, "utf-8"));
                            HtmlNodeCollection nodeInfos = docInfo.DocumentNode.SelectNodes("//div[@class='w720 fl']/table[3]/tr");
                            if (nodeInfos != null && nodeInfos.Count > 6)
                            {
                                tmpmodel.RJL = nodeInfos[1].ChildNodes[3].InnerText;
                                tmpmodel.DFL = nodeInfos[2].ChildNodes[3].InnerText;
                                tmpmodel.WYF = nodeInfos[3].ChildNodes[3].InnerText;
                                tmpmodel.LHL = nodeInfos[6].ChildNodes[3].InnerText;
                                tmpmodel.CWXX = nodeInfos[4].ChildNodes[3].InnerText;
                                tmpmodel.WYGS = nodeInfos[5].ChildNodes[3].InnerText;
                            }
                            index++;
                            Console.WriteLine("写入行：" + index);
                            tmplist.Add(tmpmodel);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.StackTrace + ex.Message);
                        }
                    }
                }
                Console.WriteLine("写入页：" + pageNum);
            }
            while (AccountToken != null);

            new ExcelOperate().DataSetToExcel(tmplist);
            Console.WriteLine("写入完毕");
            Console.ReadLine();
        }

        /// <summary>
        /// 获取网页html源文件
        /// </summary>
        /// <param name="url">网页地址</param>
        /// <param name="encodingStr">网页文件编码字符串</param>
        /// <returns>html源文件</returns>
        public static string GetPageSource(string url, string encodingStr)
        {
            HttpWebResponse res = null;
            string strResult = "";
            try
            {
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
                //req.Method = "POST";
                req.KeepAlive = true;
                req.ContentType = "application/json;charset=utf-8";
                req.Accept = "text/Html,application/xhtml+XML,application/xml;q=0.9,*/*;q=0.8";
                req.UserAgent = "Mozilla/5.0 (Windows; U; Windows NT 5.2; zh-CN; rv:1.9.2.8) Gecko/20100722 Firefox/3.6.8";
                res = (HttpWebResponse)req.GetResponse();
                StreamReader reader = new StreamReader(res.GetResponseStream(), Encoding.GetEncoding(encodingStr));
                strResult = reader.ReadToEnd();
                reader.Close();
            }
            catch (Exception ex)
            {

            }
            finally
            {
                if (res != null)
                {
                    res.Close();
                }
            }
            return strResult;
        }
    }

    /// <summary>
    /// 小区信息模类
    /// </summary>
    public class HouseInfo
    {
        public HouseInfo()
        {
            Community = "";
            Area = "";
            Position = "";
            OpenTime = "";
            Price = "";
            Measure = "";
            RJL = "";
            DFL = "";
            WYF = "";
            LHL = "";
            CWXX = "";
            WYGS = "";
        }
        /// <summary>
        /// 小区名称
        /// </summary>
        public string Community { get; set; }

        /// <summary>
        /// 区域
        /// </summary>
        public string Area { get; set; }

        /// <summary>
        /// 位置
        /// </summary>
        public string Position { get; set; }

        /// <summary>
        /// 开盘时间
        /// </summary>
        public string OpenTime { get; set; }

        /// <summary>
        /// 参考价格
        /// </summary>
        public string Price { get; set; }

        /// <summary>
        /// 户型面积
        /// </summary>
        public string Measure { get; set; }

        /// <summary>
        /// 容积率
        /// </summary>
        public string RJL { get; set; }

        /// <summary>
        /// 得房率
        /// </summary>
        public string DFL { get; set; }

        /// <summary>
        /// 物业费
        /// </summary>
        public string WYF { get; set; }

        /// <summary>
        /// 绿化率
        /// </summary>
        public string LHL { get; set; }

        /// <summary>
        /// 车位信息
        /// </summary>
        public string CWXX { get; set; }

        /// <summary>
        /// 物业公司
        /// </summary>
        public string WYGS { get; set; }
    }

    /// <summary>
    /// C#操作Excel类
    /// </summary>
    public class ExcelOperate
    {
        // 方法一：Microsoft.Office.Interop.Excel.dll（没装offic不能用，本例采用第二种）
        public bool DataSetToExcel(List<HouseInfo> dataSet, bool isShowExcle)
        {
            int rowNumber = dataSet.Count;

            if (rowNumber == 0)
            {
                return false;
            }

            //建立Excel对象
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Application.Workbooks.Add(true);
            excel.Visible = isShowExcle;//是否打开该Excel文件

            //填充数据
            for (int c = 0; c < rowNumber; c++)
            {
                excel.Cells[c + 1, 1] = dataSet[c].Community;
                excel.Cells[c + 1, 2] = dataSet[c].Area;
                excel.Cells[c + 1, 3] = dataSet[c].Price;
                excel.Cells[c + 1, 4] = dataSet[c].OpenTime;
                excel.Cells[c + 1, 5] = dataSet[c].Measure;
                excel.Cells[c + 1, 6] = dataSet[c].RJL;
                excel.Cells[c + 1, 7] = dataSet[c].DFL;
                excel.Cells[c + 1, 8] = dataSet[c].WYF;
                excel.Cells[c + 1, 9] = dataSet[c].LHL;
                excel.Cells[c + 1, 10] = dataSet[c].CWXX;
                excel.Cells[c + 1, 11] = dataSet[c].Position;
                excel.Cells[c + 1, 12] = dataSet[c].WYGS;

            }
            return true;
        }

        // 方法二：Spire.Xls.dll
        public bool DataSetToExcel(List<HouseInfo> dataSet)
        {
            try
            {
                int rowNumber = dataSet.Count;

                if (rowNumber == 0)
                {
                    return false;
                }
                Workbook book = new Workbook();
                Worksheet sheet = book.Worksheets[0];
                var random = new Random();
                var a = 0;
                //2.构造表数据
                for (int c = 0; c < rowNumber; c++)
                {
                    try
                    {
                        sheet.Range[c + 1, 1].Text = dataSet[c].Community;
                        sheet.Range[c + 1, 2].Text = dataSet[c].Area;
                        sheet.Range[c + 1, 3].Text = dataSet[c].Price;
                        sheet.Range[c + 1, 4].Text = dataSet[c].OpenTime;
                        sheet.Range[c + 1, 5].Text = dataSet[c].Measure;
                        sheet.Range[c + 1, 6].Text = dataSet[c].RJL;
                        sheet.Range[c + 1, 7].Text = dataSet[c].DFL;
                        sheet.Range[c + 1, 8].Text = dataSet[c].WYF;
                        sheet.Range[c + 1, 9].Text = dataSet[c].LHL;
                        sheet.Range[c + 1, 10].Text = dataSet[c].CWXX;
                        sheet.Range[c + 1, 11].Text = dataSet[c].Position;
                        sheet.Range[c + 1, 12].Text = dataSet[c].WYGS;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.StackTrace + ex.Message);
                    }
                }
                //3.生成图表
                book.SaveToFile(@"C:\测试\my.xlsx", ExcelVersion.Version2010);
            }
            catch (Exception ex)
            {
                Console.WriteLine("message:" + ex.Message);
            }

            return true;
        }
    }

}
