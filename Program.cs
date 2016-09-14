using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace SendEmail
{
    class Program
    {
        static void Main(string[] args)
        {
            Dns.GetHostName();
            IPAddress[] IP = Dns.GetHostAddresses(Dns.GetHostName());
            string loginIP = IP[0].ToString();
            //string dateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            //string dateTimes = DateTime.Now.ToLongTimeString();
            //SendEmailHelper.Send();
        }
        //将图片导入到excel中实例
        //public void  LoadImage()
        //{
        //    //说明：插入图片  

        //    //1.创建EXCEL中的Workbook           
        //    IWorkbook myworkbook = new HSSFWorkbook();

        //    //2.创建Workbook中的Sheet          
        //    ISheet mysheet = myworkbook.CreateSheet("sheet1");
        //    //第一步：读取图片到byte数组  
        //    HttpWebRequest request = (HttpWebRequest)WebRequest.Create("http://img1.soufunimg.com/message/images/card/tuanproj/201511/2015112703584458_s.jpg");

        //    byte[] bytes;
        //    using (Stream stream = request.GetResponse().GetResponseStream())
        //    {
        //        using (MemoryStream mstream = new MemoryStream())
        //        {
        //            int count = 0;
        //            byte[] buffer = new byte[1024];
        //            int readNum = 0;
        //            while ((readNum = stream.Read(buffer, 0, 1024)) > 0)
        //            {
        //                count = count + readNum;
        //                mstream.Write(buffer, 0, 1024);
        //            }
        //            mstream.Position = 0;
        //            using (BinaryReader br = new BinaryReader(mstream))
        //            {

        //                bytes = br.ReadBytes(count);
        //            }
        //        }
        //    }

        //    //第二步：将图片添加到workbook中  指定图片格式 返回图片所在workbook->Picture数组中的索引地址（从1开始）  
        //    int pictureIdx = myworkbook.AddPicture(bytes, PictureType.JPEG);

        //    //第三步：在sheet中创建画部  
        //    IDrawing patriarch = mysheet.CreateDrawingPatriarch();

        //    //第四步：设置锚点 （在起始单元格的X坐标0-1023，Y的坐标0-255，在终止单元格的X坐标0-1023，Y的坐标0-255，起始单元格行数，列数，终止单元格行数，列数）  
        //    IClientAnchor anchor = patriarch.CreateAnchor(0, 0, 0, 0, 0, 0, 2, 2);


        //    //第五步：创建图片  
        //    IPicture pict = patriarch.CreatePicture(anchor, pictureIdx);

        //    //6.保存         
        //    FileStream file = new FileStream(@"E:\myworkbook11.xls", FileMode.Create);
        //    myworkbook.Write(file);
        //    file.Close();
        //}
    }
}
