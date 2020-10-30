using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using System.Net.Sockets;
using System.Net;
using System.IO;
using Microsoft.Office.Core;
using System.Net.NetworkInformation;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuickDemonstrationAddIn
{
    public partial class Ribbon1
    {
        private const int BUFFER_SIZE = 1024;
        private const int PORT_NUMBER = 9999;
        static ASCIIEncoding encoding = new ASCIIEncoding();

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //InitializeComponent();
            
        }

        //Cac function cua chuong trinh:
        

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            String[] picturefile = {"https://ducbeatmusic.com/upload/posts/ojb1594810548.jpg", "https://i0.wp.com/www.tufts-skidmore.es/wp-content/uploads/2020/04/como-aprender-a-tocar-el-piano.jpg?fit=992%2C558&ssl=1"};
            Microsoft.Office.Interop.PowerPoint.Application pptapp = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation pptpre = pptapp.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
            for (int i=0; i<picturefile.Length; i++)
            {
                Microsoft.Office.Interop.PowerPoint.Slides Slides;
                Microsoft.Office.Interop.PowerPoint._Slide Slide;
                Microsoft.Office.Interop.PowerPoint.TextRange objtext;
                
                Microsoft.Office.Interop.PowerPoint.CustomLayout custlayout = pptpre.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];
                Slides = pptpre.Slides;
                Slide = Slides.AddSlide(i + 1, custlayout);
                objtext = Slide.Shapes[1].TextFrame.TextRange;
                objtext.Text = "tittle of page" + i;
                objtext.Font.Name = "Arial";
                objtext.Font.Size = 45;

                Microsoft.Office.Interop.PowerPoint.Shape shape = Slide.Shapes[2];
                Slide.Shapes.AddPicture(picturefile[i], Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, shape.Left, shape.Top, shape.Width, shape.Height);

            }

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 f = new Form1();
            if (f.ShowDialog() == DialogResult.OK)
            {
                String serverIP = f.MyVal;
                Console.WriteLine(serverIP);
                if(serverIP == "")
                {
                    MessageBox.Show("Bạn chưa điền địa chỉ");
                }
                else //connect và nhận dữ liệu
                {
                    label3.ShowLabel = false;  //chua ket noi (label 3)
                    label4.ShowLabel = true;
                    try
                    {
                       
                        TcpClient client = new TcpClient();
                        // 1. connect
                        client.Connect(serverIP, PORT_NUMBER);
                        Stream stream = client.GetStream();
                        Console.WriteLine("Connected to Y2Server.");
                        MessageBox.Show("Đã kết nối đến điện thoại");
                        label3.ShowLabel = false;
                        label4.ShowLabel = false;
                        label2.ShowLabel = true;
                        Boolean dk = true;

                        Microsoft.Office.Interop.PowerPoint._Application pApp = new Microsoft.Office.Interop.PowerPoint.Application();
                        Microsoft.Office.Interop.PowerPoint.Presentation pRe = pApp.Presentations.Open(pApp.ActivePresentation.FullName, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                        Microsoft.Office.Interop.PowerPoint.Slides Slides;
                        Microsoft.Office.Interop.PowerPoint._Slide Slide;
                        Microsoft.Office.Interop.PowerPoint.TextRange objtext;
                        Microsoft.Office.Interop.PowerPoint.CustomLayout custlayout = pRe.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];
                        
                        while (dk == true)
                        {
                            //Console.Write("Enter your name: ");

                            //string str = Console.ReadLine();
                            var reader = new StreamReader(stream);
                            var writer = new StreamWriter(stream);
                            writer.AutoFlush = true;

                            // 2. send
                            byte[] data = encoding.GetBytes("Data from laptop (byte)");
                            //stream.Write(data, 0, data.Length);
                            //writer.WriteLine("Data from laptop");

                            // 3. receive
                            data = new byte[BUFFER_SIZE];
                            stream.Read(data, 0, BUFFER_SIZE);
                            String data_recieve = encoding.GetString(data);
                            //MessageBox.Show(data_recieve);

                            String str = reader.ReadLine();
                            //MessageBox.Show(str);  //string sau
                            //MessageBox.Show(data_recieve);  //ki tu dau
                            //Console.WriteLine("Command from phone 2: " + str);
                            //MessageBox.Show("Command from phone 2: " + str);

                            //Code chinh cho việc điều khiển 
                            if (String.Compare(data_recieve, "next") == 0)
                            {
                                //MessageBox.Show("Di chuyen den slide tiep theo . . .");
                                pRe.SlideShowWindow.View.Next();
                            }
                            else if (String.Compare(data_recieve, "pre") == 0)
                            {
                                //MessageBox.Show("Di chuyen den slide trước đó . . .");
                                pRe.SlideShowWindow.View.Previous();
                            }
                            else if (String.Compare(data_recieve+str, "start") == 0)
                            {
                                try
                                {
                                    //MessageBox.Show("Bắt đầu trình chiếu . . .");
                                    //MessageBox.Show(pApp.ActivePresentation.FullName);
                                    pRe.SlideShowSettings.Run();
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.ToString());
                                }
                            }
                            else if (String.Compare(data_recieve, "end") == 0)
                            {
                                //MessageBox.Show("Kết thúc trình chiếu . . .");
                                pRe.SlideShowWindow.View.Exit();
                            }
                            else if (String.Compare(data_recieve, "startS") == 0)
                            {
                                MessageBox.Show("Di chuyen den slide đầu tiên . . .");
                            }
                            else if (String.Compare(data_recieve, "endS") == 0)
                            {
                                MessageBox.Show("Di chuyen den slide cuối cùng . . .");
                            }

                            else if (String.Compare(data_recieve.ToUpper(), "BYE") == 0)   //Đóng kết nối
                            {
                                label3.ShowLabel = true;
                                label2.ShowLabel = false;
                                dk = false;
                                break;
                            }

                            else
                            {
                                try
                                {
                                    MessageBox.Show(data_recieve);
                                    /*Slides = pRe.Slides;
                                    Slide = Slides.AddSlide(pRe.Slides.Count + 1, custlayout);
                                    objtext = Slide.Shapes[1].TextFrame.TextRange;
                                    objtext.Text = data_recieve;
                                    objtext.Font.Name = "Arial";
                                    objtext.Font.Size = 45;

                                    Microsoft.Office.Interop.PowerPoint.Shape shape = Slide.Shapes[2];
                                    Slide.Shapes.AddPicture(data_recieve, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, shape.Left, shape.Top, shape.Width, shape.Height);
                                    //pRe.Slides[pRe.Slides.Count].Select();
                                    //Slide = pRe.Slides[pRe.Slides.Count];
                                    pRe.SlideShowSettings.Run();
                                    pRe.SlideShowWindow.View.Last();*/
                                }
                                catch(Exception ex)
                                {
                                    MessageBox.Show(ex.ToString());
                                }
                                break;
                            }
                        }
                        
                        // 4. close
                        stream.Close();
                        client.Close();
                        MessageBox.Show("Đã ngắt kết nối với điện thoại!");
                        label3.ShowLabel = true;
                        label2.ShowLabel = false;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error: " + ex);
                        MessageBox.Show(ex.ToString());
                        label4.ShowLabel = false;
                        label3.ShowLabel = true;
                        label2.ShowLabel = false;

                    }
                }
            }
            else 
            {
                MessageBox.Show("Kết nối chưa được thiết lập, vui lòng thiết lập lại địa chỉ kết nối!");
            }
            //String picturefile = "https://ducbeatmusic.com/upload/posts/ojb1594810548.jpg";
        }


        //Phần kết nối chính với điện thoại (PC là server)
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            QRform qr = new QRform();
            qr.Show();
        }
    }
}
