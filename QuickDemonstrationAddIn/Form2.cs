using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media.Animation;

namespace QuickDemonstrationAddIn
{
    public partial class QRform : Form
    {
        
        public QRform()
        {
            InitializeComponent();
        }

        //Cac function chính của chương trình:
        private String getIP(NetworkInterfaceType _type)  //Lấy địa chỉ IP của PC
        {
            string output = "";
            foreach (NetworkInterface item in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (item.NetworkInterfaceType == _type && item.OperationalStatus == OperationalStatus.Up)
                {
                    foreach (UnicastIPAddressInformation ip in item.GetIPProperties().UnicastAddresses)
                    {
                        if (ip.Address.AddressFamily == AddressFamily.InterNetwork)
                        {
                            output = ip.Address.ToString();
                        }
                    }
                }
            }
            return output;
        }

        private void make_qr() //Tạo mã QR
        {
            QRCoder.QRCodeGenerator QG = new QRCoder.QRCodeGenerator();
            var MyData = QG.CreateQrCode(getIP(NetworkInterfaceType.Wireless80211), QRCoder.QRCodeGenerator.ECCLevel.H);
            var code = new QRCoder.QRCode(MyData);
            pictureBox1.Image = code.GetGraphic(9);
        }

        private void note()
        {
            try
            {
                Microsoft.Office.Interop.PowerPoint._Application pApp = new Microsoft.Office.Interop.PowerPoint.Application();
                Microsoft.Office.Interop.PowerPoint.Presentation pRe = pApp.Presentations.Open(pApp.ActivePresentation.FullName, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                Microsoft.Office.Interop.PowerPoint.Slide slide = pRe.SlideShowWindow.View.Slide;
                MessageBox.Show(pRe.SlideShowWindow.View.Slide.NotesPage.Shapes[0].TextFrame.TextRange.Text);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
            }
            
        }

        //Phần code chạy server socket
        private void server()
        {

            QRform qr = new QRform();
            bool dk = true;
            try
            {
                IPAddress ipAd = IPAddress.Parse(getIP(NetworkInterfaceType.Wireless80211));
                TcpListener myList = new TcpListener(ipAd, 9999);
                myList.Start();
                Socket s = myList.AcceptSocket();
                hideForm(true);
                MessageBox.Show("Đã kết nối với điện thoại :)");
                ASCIIEncoding asen = new ASCIIEncoding();
                try
                {
                   
                    Microsoft.Office.Interop.PowerPoint._Application pApp = new Microsoft.Office.Interop.PowerPoint.Application();
                    Microsoft.Office.Interop.PowerPoint.Presentation pRe = pApp.Presentations.Open(pApp.ActivePresentation.FullName, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                    Microsoft.Office.Interop.PowerPoint.SlideShowWindow ssw;
                    Microsoft.Office.Interop.PowerPoint.Slides Slides;
                    Microsoft.Office.Interop.PowerPoint._Slide Slide;
                    Microsoft.Office.Interop.PowerPoint.TextRange objtext;
                    Microsoft.Office.Interop.PowerPoint.CustomLayout custlayout = pRe.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];

                    while (dk==true)
                    {
                        byte[] b = new byte[1000];
                        int k = s.Receive(b);
                        char cc = ' ';
                        string command = "";
                        for (int i = 0; i < k; i++)
                        {
                            cc = Convert.ToChar(b[i]);
                            command += cc.ToString();
                        }
                        //MessageBox.Show(command);

                        //Codde lenh dieu khien 
                        //Code chinh cho việc điều khiển 
                        if (command.Contains("disconnect"))
                        {
                            s.Close();
                            myList.Stop();
                            MessageBox.Show("Đã ngắt kết nối với điện thoại!");
                            hideForm(false);

                        }
                        if (command.Contains("next"))
                        {
                            //MessageBox.Show("Di chuyen den slide tiep theo . . .");
                            pRe.SlideShowWindow.View.Next();
                            ssw = pRe.SlideShowWindow.Presentation.SlideShowWindow;
                            if (ssw.View.Slide.HasNotesPage == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                if (ssw.View.Slide.NotesPage.Shapes[2].TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                                    MessageBox.Show(ssw.View.Slide.NotesPage.Shapes[2].TextFrame.TextRange.Text);
                            }
                        }
                    
                        if (command.Contains("pre"))
                        {
                            pRe.SlideShowWindow.View.Previous();
                            ssw = pRe.SlideShowWindow.Presentation.SlideShowWindow;
                            if (ssw.View.Slide.HasNotesPage == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                if (ssw.View.Slide.NotesPage.Shapes[2].TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                                    MessageBox.Show(ssw.View.Slide.NotesPage.Shapes[2].TextFrame.TextRange.Text);
                            }
                        }
                        if (command.Contains("start"))
                        {
                            pApp = new Microsoft.Office.Interop.PowerPoint.Application();
                            pRe = pApp.Presentations.Open(pApp.ActivePresentation.FullName, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                            pRe.SlideShowSettings.Run();
                            ssw = pRe.SlideShowWindow.Presentation.SlideShowWindow;
                            if (ssw.View.Slide.HasNotesPage == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                if (ssw.View.Slide.NotesPage.Shapes[2].TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                                    s.Send(asen.GetBytes(ssw.View.Slide.NotesPage.Shapes[2].TextFrame.TextRange.Text));  //gửi data note đến Android
                                    //MessageBox.Show(ssw.View.Slide.NotesPage.Shapes[2].TextFrame.TextRange.Text);
                            }
                            //Thread noteThread = new Thread(new ThreadStart(note));
                            //noteThread.Start();
                            
                        }
                        if (command.Contains("end"))
                        {
                            pRe.SlideShowWindow.View.Exit();
                        }
                        if (command.Contains("first"))
                        {
                            pRe.SlideShowWindow.View.First();
                            ssw = pRe.SlideShowWindow.Presentation.SlideShowWindow;
                            if (ssw.View.Slide.HasNotesPage == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                if (ssw.View.Slide.NotesPage.Shapes[2].TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                                    MessageBox.Show(ssw.View.Slide.NotesPage.Shapes[2].TextFrame.TextRange.Text);
                            }
                        }
                        if (command.Contains("last"))
                        {
                            pRe.SlideShowWindow.View.Last();
                            ssw = pRe.SlideShowWindow.Presentation.SlideShowWindow;
                            if (ssw.View.Slide.HasNotesPage == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                if (ssw.View.Slide.NotesPage.Shapes[2].TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                                    MessageBox.Show(ssw.View.Slide.NotesPage.Shapes[2].TextFrame.TextRange.Text);
                            }
                        }
                        if (command.Contains("click"))
                        {
                            MessageBox.Show(pRe.SlideShowWindow.View.Slide.NotesPage.ToString());
                        }

                        //Phần chèn ảnh và text
                        if (command.Contains("http"))
                        {
                            //MessageBox.Show(command);
                            try
                            {
                                //pApp = new Microsoft.Office.Interop.PowerPoint.Application();
                                //pRe = pApp.Presentations.Open(pApp.ActivePresentation.FullName, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                                Slides = pRe.Slides;
                                Slide = Slides.AddSlide(pRe.Slides.Count + 1, custlayout);
                                //objtext = Slide.Shapes[1].TextFrame.TextRange;
                                //objtext.Text = "image";
                                //objtext.Font.Name = "Arial";
                                //objtext.Font.Size = 45;

                                //Microsoft.Office.Interop.PowerPoint.Shape shape = Slide.Shapes[1];
                                //Slide.Shapes.AddPicture(command, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, shape.Left, shape.Top, shape.Width, shape.Height);

                                var slidelast = pRe.Slides[pRe.Slides.Count];
                                slidelast.FollowMasterBackground = MsoTriState.msoFalse;
                                slidelast.Background.Fill.UserPicture(command);

                                slidelast.Select();
                                Slide = pRe.Slides[pRe.Slides.Count];
                                pRe.SlideShowSettings.Run();
                                pRe.SlideShowWindow.View.Last();
                            }
                            catch (Exception ex)
                            {
                                pRe.SlideShowWindow.View.Last();
                                //MessageBox.Show(ex.ToString());
                            }
                        }
                        


                        //ASCIIEncoding asen = new ASCIIEncoding();
                        //s.Send(asen.GetBytes("The string send from the PC."));  //gửi data đến Android
                    }
                    s.Close();
                    myList.Stop();
                    hideForm(false);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message.ToString());
                    //MessageBox.Show("Bạn chưa mở file Powerpoint nên chưa thể dùng tình năng này!\nVui lòng mở một file cụ thể trước rồi thử kết nối lại với điện thoại!");
                    s.Close();
                    myList.Stop();
                    hideForm(false);
                }
            }
            catch (Exception e)
            {
                //MessageBox.Show("Kết nối chưa được thiết lập, vui lòng thiết lập lại địa chỉ kết nối!");
                MessageBox.Show(e.ToString());
                hideForm(false);
            }
        }

        private delegate void SafeCallDelegate(bool b);
        private void hideForm(bool b)
        {
            if (this.InvokeRequired)
            {
                if (!b)
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        this.Visible = false;
                    }));
                }
                else
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        this.Visible = true;
                        this.Close();
                    }));
                }
            }
            else
            {
                if (!b)
                {
                    this.Visible = false;
                }
                else
                {
                    this.Visible = true;
                    this.Close();
                }
            }
        }

        private void QRform_Load(object sender, EventArgs e)
        {
            Thread serverThread = new Thread(new ThreadStart(server));
            serverThread.Start();
            make_qr();
            label2.Text = getIP(NetworkInterfaceType.Wireless80211);
        }

        private void button1_Click(object sender, EventArgs e)  //Nút Hủy
        {
            Thread serverThread = new Thread(new ThreadStart(server));
            serverThread.Abort();
            this.Close();
        }
    }
}
