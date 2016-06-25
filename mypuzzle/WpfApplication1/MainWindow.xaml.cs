using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
//add for mouse
using System.Windows.Threading;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;


namespace WpfApplication1
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
           // timer顯示座標
            DispatcherTimer timersec,timershow;
            MouseButtonEventArgs curmouse = null;
            bool changemode = false;             //手動變換珠子顏色 判斷是不是轉珠模式
            int wood = 0, water = 0, fire = 0, gold = 0, dark = 0, heart = 0;
            int comwood = 0, comwater = 0, comfire = 0, comgold = 0, comdark = 0, comheart = 0;
            int countwood = 0, countwater = 0, countfire = 0, countgold = 0, countdark = 0, countheart = 0;
            int[,] orb = new int[5,6];
            int[,] zero = new int[5, 6] { { 0, 0, 0, 0, 0, 0 }, { 0, 0, 0, 0, 0, 0 }, { 0, 0, 0, 0, 0, 0 }, { 0, 0, 0, 0, 0, 0 }, { 0, 0, 0, 0, 0, 0 } };

        public MainWindow()
        {
              // timer顯示座標
            DispatcherTimer Postiontimer;
            Postiontimer = new DispatcherTimer();
            Postiontimer.Tick += Postiontimer_1Tick;
            Postiontimer.Start();
            //          
            InitializeComponent();
            creatpuzzle(5,6);               
        }
        
        private void creatpuzzle(int x, int y) //動態生成珠子
        {          
            wood = 0; water = 0; fire = 0; gold = 0; dark = 0; heart = 0;
            for (int i = 0; i < x; i++) {
                for (int j = 0; j < y; j++) {
                    System.Windows.Shapes.Rectangle rec = new System.Windows.Shapes.Rectangle()
                    {
                        Height = 50,
                        Width = 50,
                        //Stroke = new SolidColorBrush(Colors.Transparent),                      
                    };
                    randomfill(i , j , rec);    //隨機給珠子顏色           
                    rec.MouseDown +=  new MouseButtonEventHandler(this.Rectangle_MouseDown); 
                    rec.MouseMove +=  new MouseEventHandler(this.Rectangle_MouseMove);
                    rec.MouseUp += new MouseButtonEventHandler(this.Rectangle_MouseUp);
                    Canvas.SetTop(rec, 40 +i*55);
                    Canvas.SetLeft(rec, 40+ j*55);                    
                    can1.Children.Add(rec);
                }
            }
            lbdark.Content = "暗珠: " + dark + "顆";
            lbwater.Content = "水珠: " + water + "顆";
            lbfire.Content = "火珠: " + fire + "顆";
            lbwood.Content = "木珠: " + wood + "顆";
            lbgold.Content = "光珠: " + gold + "顆";
            lbheart.Content = "心珠: " + heart + "顆";
            lbcombo.Content = "最大combo數: " + ((dark / 3) + (water / 3) + (fire / 3) + (wood / 3) + (gold / 3) + (heart / 3));
        }

        private void randomfill(int a , int b , System.Windows.Shapes.Rectangle rec)//隨機上色
        {
            Random rnd = new Random(Guid.NewGuid().GetHashCode());
            int x = rnd.Next(6);
            ImageBrush brush = null;
            //問助教為什麼要用image 不能用new的   ImageBrush brush = new ImageBrush(new BitmapImage(new Uri("images/gold.png", UriKind.Relative)));      
            switch (x)
            {
                case 1:
                    image1.Source = new BitmapImage(new Uri("images/gold.png", UriKind.Relative));
                    brush = new ImageBrush(image1.Source);
                    gold++;
                    orb[a, b] = 1;
                    break;
                case 2:
                    image1.Source = new BitmapImage(new Uri("images/dark.png", UriKind.Relative));
                    brush = new ImageBrush(image1.Source);
                    dark++;
                    orb[a, b] = 2;
                    break;
                case 3:
                    image1.Source = new BitmapImage(new Uri("images/fire.png", UriKind.Relative));
                    brush = new ImageBrush(image1.Source);
                    fire++;
                    orb[a, b] = 3;
                    break;
                case 4:
                    image1.Source = new BitmapImage(new Uri("images/water.png", UriKind.Relative));
                    brush = new ImageBrush(image1.Source);
                    water++;
                    orb[a, b] = 4;
                    break;
                case 5:
                    image1.Source = new BitmapImage(new Uri("images/wood.png", UriKind.Relative));
                    brush = new ImageBrush(image1.Source);
                    wood++;
                    orb[a, b] = 5;
                    break;
                case 0:
                    image1.Source = new BitmapImage(new Uri("images/heart.png", UriKind.Relative));
                    brush = new ImageBrush(image1.Source);
                    heart++;
                    orb[a, b] = 6;
                    break;
                default:
                    MessageBox.Show("no puzzle!!!!!");
                    break;
            }
            //   ImageBrush brush = new ImageBrush(new BitmapImage(new Uri("picture/gold.png", UriKind.Relative)));
            rec.Fill = brush;
            return;
        }

        double mouseX;
        double mouseY;
        double oldpositionX;    //圖形原本位置
        double oldpositionY;    
        double currentshapX;    //圖形隨著滑鼠移動的位置
        double currentshapY;
        System.Windows.Shapes.Rectangle CurrentRec = null;  //宣告目前移動的正方形

        private int find_combo(int[,] ori)
        {
            //int[,] ori = orb;
            int count = 0;
            bool judge = false;
            int[,] temp = new int[5 , 6];
            // 1. filter all 3+ consecutives.
            //  (a) horizontals
            for (var i = 0; i < 5; ++ i) {
                var prev_1_orb = 19;
                var prev_2_orb = 19;
                for (var j = 0; j < 6; ++ j) {
                    var cur_orb = ori[i , j];
                    if (prev_1_orb == prev_2_orb && prev_2_orb == cur_orb && cur_orb != 19 && cur_orb != 29) {
                        temp[i , j] = cur_orb;
                        temp[i , j-1] = cur_orb;
                        temp[i , j-2] = cur_orb;
                    }
                    prev_1_orb = prev_2_orb;
                    prev_2_orb = cur_orb;
                }
            }
            //  (b) verticals
            for (var j = 0; j < 6; ++ j) {
                var prev_1_orb = 19;
                var prev_2_orb = 19;
                for (var i = 0; i < 5; ++ i) {
                    var cur_orb = ori[i , j];
                    if (prev_1_orb == prev_2_orb && prev_2_orb == cur_orb && cur_orb != 19 && cur_orb != 29) {
                        temp[i , j] = cur_orb;
                        temp[i-1 , j] = cur_orb;
                        temp[i-2 , j] = cur_orb;
                    }
                    prev_1_orb = prev_2_orb;
                    prev_2_orb = cur_orb;
                }
            }

            int[,] temp2 = ori;
            /*for (int i = 0; i < 5; i++)
            {
                for (int j = 0; j < 6; j++)
                    temp2[i, j] = orb[i, j];
                {
                    textbox1.Text += temp[i, j].ToString() + " ";
                }
                textbox1.Text += "\n";
            }*/
            for (int i = 0; i < 5; i++)
            {
                for (int j = 0; j < 6; j++)
                {
                    if (temp2[i, j] == temp[i, j])
                    {
                        judge = true;
                        temp2[i, j] = 29;
                        for (int k = i; k > 0; k--)
                        {
                            int oritemp = temp2[k, j];
                            temp2[k, j] = temp2[k - 1, j];
                            temp2[k - 1, j] = oritemp;
                        }
                    }
                }
            }
            for (int i = 0; i < 5; i++)
            {
                for (int j = 0; j < 6; j++)
                {
                    //textbox1.Text += temp[i, j].ToString() + " ";
                }
                //textbox1.Text += "\n";
            }
            //textbox1.Text += "\n";
            count = combocount(temp);
            if (judge == true)
                return find_combo(temp2) + count;
            else
                return 0;
        }

        private int combocount(int[,] ori)
        {
            //int[,] temp = new int[6, 7];
            //bool combojudge = false;
            int count = 0;
            /*for (int i = 0; i < 5; i++)
                for (int j = 0; j < 6; j++)
                    temp[i + 1, j + 1] = ori[i, j];*/
            for (int i = 0; i < 5; i++)
                for (int j = 0; j < 6; j++)
                {
                    if (ori[i , j] != 0)
                    {
                        switch (ori[i, j])
                        {
                            case 1:
                                comgold++;
                                break;
                            case 2:
                                comdark++;
                                break;
                            case 3:
                                comfire++;
                                break;
                            case 4:
                                comwater++;
                                break;
                            case 5:
                                comwood++;
                                break;
                            case 6:
                                comheart++;
                                break;
                        }
                        //combojudge = true;
                        ori = setzero(ori , i , j);
                        count++;
                    }
                }
            return count;
        }

        private int[,] setzero(int[,] ori , int i , int j)
        {
            int temp = ori[i, j];
            switch (ori[i, j])
            {
                case 1:
                    countgold++;
                    break;
                case 2:
                    countdark++;
                    break;
                case 3:
                    countfire++;
                    break;
                case 4:
                    countwater++;
                    break;
                case 5:
                    countwood++;
                    break;
                case 6:
                    countheart++;
                    break;
            }
            ori[i, j] = 0;
            if (i > 0)
                if (temp == ori[i - 1, j] )
                    ori =  setzero(ori, i - 1, j);
            if (i < 4)
                if (temp  == ori[i + 1, j])
                    ori =  setzero(ori, i + 1, j);
            if (j > 0)
                if (temp  == ori[i, j - 1])
                    ori =  setzero(ori, i, j - 1);
            if (j < 5)
                if (temp  == ori[i, j+1])
                    ori =  setzero(ori, i, j + 1);
            return ori;
        }
                                               
        private void Rectangle_MouseDown(object sender, MouseButtonEventArgs e)
        {

            if (changemode)  //換色模式
            {
                double x, y;
                System.Windows.Shapes.Rectangle item = sender as System.Windows.Shapes.Rectangle;
               // item.Fill = new SolidColorBrush(Colors.Red);  //delete later, just for mouse test
                ImageBrush brush = null;
                //image1.Source = new BitmapImage(new Uri("images/dark.png", UriKind.Relative));
                //brush = new ImageBrush(image1.Source);
                //item.Fill = brush;
                y = (double)item.GetValue(Canvas.LeftProperty);
                x = (double)item.GetValue(Canvas.TopProperty);
                x = (x - 40) / 55 ;
                y = (y - 40) / 55 ;
                lb1.Content = x;
                lb2.Content = y;
                if (zero[(int)x, (int)y] == 0)
                {
                    zero[(int)x, (int)y] = 1;
                    orb[(int)x, (int)y] = 6;
                }
                switch (orb[(int)x, (int)y])
                {
                    case 6:
                        image1.Source = new BitmapImage(new Uri("images/gold.png", UriKind.Relative));
                        brush = new ImageBrush(image1.Source);
                        item.Fill = brush;
                        heart--;
                        gold++;
                        orb[(int)x, (int)y] = 1;
                        break;
                    case 1:
                        image1.Source = new BitmapImage(new Uri("images/dark.png", UriKind.Relative));
                        brush = new ImageBrush(image1.Source);
                        item.Fill = brush;
                        gold--;
                        dark++;
                        orb[(int)x, (int)y] = 2;
                        break;
                    case 2:
                        image1.Source = new BitmapImage(new Uri("images/fire.png", UriKind.Relative));
                        brush = new ImageBrush(image1.Source);
                        item.Fill = brush;
                        dark--;
                        fire++;
                        orb[(int)x, (int)y] = 3;
                        break;
                    case 3:
                        image1.Source = new BitmapImage(new Uri("images/water.png", UriKind.Relative));
                        brush = new ImageBrush(image1.Source);
                        item.Fill = brush;
                        fire--;
                        water++;
                        orb[(int)x, (int)y] = 4;
                        break;
                    case 4:
                        image1.Source = new BitmapImage(new Uri("images/wood.png", UriKind.Relative));
                        brush = new ImageBrush(image1.Source);
                        item.Fill = brush;
                        water--;
                        wood++;
                        orb[(int)x, (int)y] = 5;
                        break;
                    case 5:
                        image1.Source = new BitmapImage(new Uri("images/heart.png", UriKind.Relative));
                        brush = new ImageBrush(image1.Source);
                        item.Fill = brush;
                        wood--;
                        heart++;
                        orb[(int)x, (int)y] = 6;
                        break;
                    default:
                        MessageBox.Show("no puzzle!!!!!");
                        break;
                }
                lbdark.Content = "暗珠: " + dark + "顆";
            lbwater.Content = "水珠: " + water + "顆";
            lbfire.Content = "火珠: " + fire + "顆";
            lbwood.Content = "木珠: " + wood + "顆";
            lbgold.Content = "光珠: " + gold + "顆";
            lbheart.Content = "心珠: " + heart + "顆";
            lbcombo.Content = "最大combo數: " + ((dark / 3) + (water / 3) + (fire / 3) + (wood / 3) + (gold / 3) + (heart / 3));
            }
            else  //轉珠模式
            {
                //timer
                timersec = new DispatcherTimer();
                timersec.Interval = TimeSpan.FromSeconds(5.0);
                timersec.Tick += timersec_Tick;
                timersec.Start();
                //
                timershow = new DispatcherTimer();
                timershow.Interval = TimeSpan.FromSeconds(1.0);
                timershow.Tick += timershow_Tick;
                this.lbtimer.Content = "0";
                timershow.Start();
                //

                curmouse = e;

                System.Windows.Shapes.Rectangle item = sender as System.Windows.Shapes.Rectangle;
                CurrentRec = item;  //記錄下要移動的正方形(hittest才可以判斷)
                // item.Fill = new SolidColorBrush(Colors.Red);  //delete later, just for mouse test
                Canvas.SetZIndex(item, 100);  //移動的物體總是在最上方
                mouseX = e.GetPosition(null).X;
                mouseY = e.GetPosition(null).Y;
                item.CaptureMouse();
                oldpositionX = (double)item.GetValue(Canvas.LeftProperty);
                oldpositionY = (double)item.GetValue(Canvas.TopProperty);
                lb1.Content = oldpositionX;
                lb2.Content = oldpositionY;
            }
            
        }

        private void Rectangle_MouseUp(object sender, MouseButtonEventArgs e)
        {
            System.Windows.Shapes.Rectangle item = sender as System.Windows.Shapes.Rectangle;
            if (!changemode && item.IsMouseCaptured)  //如果是上色模式就不執行
            {
                
                timershow.Stop();
                timersec.Stop();
                item.ReleaseMouseCapture();
                Canvas.SetZIndex(item, 0);  //物體位置回歸水平
                mouseX = -1;
                mouseY = -1;
                item.SetValue(Canvas.LeftProperty, oldpositionX); //轉完自動歸位圓圈
                item.SetValue(Canvas.TopProperty, oldpositionY);  //  
            }
        }

        private void Rectangle_MouseMove(object sender, MouseEventArgs e)
        {                       
            System.Windows.Shapes.Rectangle item = sender as System.Windows.Shapes.Rectangle;
            if (item.IsMouseCaptured)
            {
                // Calculate the current position of the object.
                double deltaX = e.GetPosition(null).X - mouseX;
                double deltaY = e.GetPosition(null).Y - mouseY;
                double newLeft = deltaX + (double)item.GetValue(Canvas.LeftProperty);
                double newTop = deltaY + (double)item.GetValue(Canvas.TopProperty);
                // Set new position of object.
                item.SetValue(Canvas.LeftProperty, newLeft);
                item.SetValue(Canvas.TopProperty, newTop);
                //add by me
                currentshapX = (double)item.GetValue(Canvas.LeftProperty);
                currentshapY = (double)item.GetValue(Canvas.TopProperty);
                // Update position global variables.
                mouseX = e.GetPosition(null).X;
                mouseY = e.GetPosition(null).Y;
              //  將正方形形圖轉成幾何圖形
                Geometry g = item.RenderedGeometry;
            //    座標位置轉換為視窗的座標
                g.Transform = item.TransformToAncestor(this) as MatrixTransform;
                VisualTreeHelper.HitTest(this, null,
                    new HitTestResultCallback(myHitTestResult),
                    new GeometryHitTestParameters(g));
            }
        }
           
  
        public HitTestResultBehavior myHitTestResult(HitTestResult result)
        {
            if (result.VisualHit is System.Windows.Shapes.Rectangle && result.VisualHit!= CurrentRec)
            {
                System.Windows.Shapes.Rectangle rect = result.VisualHit as System.Windows.Shapes.Rectangle;
            //    rect.Fill = new SolidColorBrush(Colors.Red);            
                //if (currentshapX > oldpositionX + 25 || currentshapX < oldpositionX - 25 || currentshapY > oldpositionY + 25 || currentshapY < oldpositionY - 25)
                if (((currentshapX - oldpositionX) * (currentshapX - oldpositionX) + (currentshapY - oldpositionY) * (currentshapY - oldpositionY)) > 1200)
                {
                    double tempX=0, tempY=0;
                    int temp;
                    tempX = (double)rect.GetValue(Canvas.LeftProperty);                  
                    tempY = (double)rect.GetValue(Canvas.TopProperty);
                    rect.SetValue(Canvas.LeftProperty, oldpositionX);
                    rect.SetValue(Canvas.TopProperty, oldpositionY);
                    lb1.Content = currentshapX;
                    lb2.Content = currentshapY;
                    temp = orb[(int)(oldpositionY - 40) / 55, (int)(oldpositionX - 40) / 55];
                    orb[(int)(oldpositionY - 40) / 55, (int)(oldpositionX - 40) / 55] = orb[(int)(tempY - 40) / 55, (int)(tempX - 40) / 55];
                    orb[(int)(tempY - 40) / 55, (int)(tempX - 40) / 55] = temp;
                    oldpositionX = tempX;
                    oldpositionY = tempY;
                    lb1.Content = oldpositionX;
                    lb2.Content = oldpositionY;
                }            
            }
            return HitTestResultBehavior.Continue;
        }

        private void btrand_Click(object sender, RoutedEventArgs e)
        {
            can1.Children.RemoveRange(18,35); //從第18個產生的物件開始移除(移除拼圖)
            creatpuzzle(5, 6);
        }

        private void btmove_Click(object sender, RoutedEventArgs e)
        {           
           // Mouse.MoveTo(tbx.Text, tby.Text);
           // Mouse.LeftClick();           
        }

        private void btchange_color_Click(object sender, RoutedEventArgs e)
        {
            changemode = true;
        }

        private void btspin_mode_Click(object sender, RoutedEventArgs e)
        {
            changemode = false;
        }

        void timersec_Tick(object sender, EventArgs e)
        {
            // this.lbtimer.Content = (int.Parse(this.lbtimer.Content.ToString()) + 1).ToString();
            timersec.Stop();
            timershow.Stop();
            this.lbtimer.Content += "秒到了";
            Rectangle_MouseUp(CurrentRec, curmouse);
        }

        private void timershow_Tick(object sender, EventArgs e)
        {
            this.lbtimer.Content = (int.Parse(this.lbtimer.Content.ToString()) + 1).ToString();
            if (int.Parse(this.lbtimer.Content.ToString()) == 5)
            {
                timershow.Stop();
                this.lbtimer.Content += "秒到了";

            }
        }

        private void Postiontimer_1Tick(object sender, EventArgs e)
        {
            this.plabel.Content = "(" + System.Windows.Forms.Cursor.Position.X.ToString() + "," + System.Windows.Forms.Cursor.Position.Y.ToString() + ")";
        }

        private void btn_combo_Click(object sender, RoutedEventArgs e)
        {
            comwood = 0; comwater = 0; comfire = 0; comgold = 0; comdark = 0; comheart = 0;
            countwood = 0; countwater = 0; countfire = 0; countgold = 0; countdark = 0; countheart = 0;
            int count;
            int[,] temp = new int[5,6];
            for (int i = 0; i < 5; i++)
            {
                for (int j = 0; j < 6; j++)
                {
                    temp[i, j] = orb[i, j];
                    //textbox1.Text += orb[i, j].ToString() + " ";
                }
                //textbox1.Text += "\n";
            }
            count = find_combo(temp);
            textbox1.Text = "總共有 " + count.ToString() + " Combo\n";
            tbdark.Text = "總共 " + comdark + " Combo 共 " + countdark + "顆";
            tbgold.Text = "總共 " + comgold + " Combo 共 " + countgold + "顆";
            tbwood.Text = "總共 " + comwood + " Combo 共 " + countwood + "顆";
            tbwater.Text = "總共 " + comwater + " Combo 共 " + countwater + "顆";
            tbfire.Text = "總共 " + comfire + " Combo 共 " + countfire + "顆";
            tbheart.Text = "總共 " + comheart + " Combo 共 " + countheart + "顆";
        }
    }

    //add for mouse
    internal class NativeContansts
    {
        public const int WH_MOUSE_LL = 14;
        public const int WH_KEYBOARD_LL = 13;
        public const int WH_MOUSE = 7;
        public const int WH_KEYBOARD = 2;

        public const int WM_MOUSEMOVE = 0x200;
        public const int WM_LBUTTONDOWN = 0x201;
        public const int WM_RBUTTONDOWN = 0x204;
        public const int WM_MBUTTONDOWN = 0x207;
        public const int WM_LBUTTONUP = 0x202;
        public const int WM_RBUTTONUP = 0x205;
        public const int WM_MBUTTONUP = 0x208;
        public const int WM_LBUTTONDBLCLK = 0x203;
        public const int WM_RBUTTONDBLCLK = 0x206;
        public const int WM_MBUTTONDBLCLK = 0x209;
        public const int WM_MOUSEWHEEL = 0x020A;
        public const int WM_KEYDOWN = 0x100;
        public const int WM_KEYUP = 0x101;
        public const int WM_SYSKEYDOWN = 0x104;
        public const int WM_SYSKEYUP = 0x105;

        public const int MEF_LEFTDOWN = 0x00000002;
        public const int MEF_LEFTUP = 0x00000004;
        public const int MEF_MIDDLEDOWN = 0x00000020;
        public const int MEF_MIDDLEUP = 0x00000040;
        public const int MEF_RIGHTDOWN = 0x00000008;
        public const int MEF_RIGHTUP = 0x00000010;

        public const int KEF_EXTENDEDKEY = 0x1;
        public const int KEF_KEYUP = 0x2;

        public const byte VK_SHIFT = 0x10;
        public const byte VK_CAPITAL = 0x14;
        public const byte VK_NUMLOCK = 0x90;

        public const int WM_IME_SETCONTEXT = 0x0281;
        public const int WM_CHAR = 0x0102;
        public const int WM_IME_COMPOSITION = 0x010F;
        public const int GCS_COMPSTR = 0x0008;
    }

    public static class Mouse
    {
        [StructLayout(LayoutKind.Explicit, Pack = 1, Size = 28)]
        public struct INPUT
        {
            [FieldOffset(0)]
            public Mouse.INPUTTYPE dwType;
            [FieldOffset(4)]
            public Mouse.MOUSEINPUT mi;
            [FieldOffset(4)]
            public Mouse.KEYBOARDINPUT ki;
            [FieldOffset(4)]
            public Mouse.HARDWAREINPUT hi;
        }
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public int mouseData;
            public Mouse.MOUSEFLAG dwFlags;
            public int time;
            public IntPtr dwExtraInfo;
        }
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct KEYBOARDINPUT
        {
            public short wVk;
            public short wScan;
            public Mouse.KEYBOARDFLAG dwFlags;
            public int time;
            public IntPtr dwExtraInfo;
        }
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct HARDWAREINPUT
        {
            public int uMsg;
            public short wParamL;
            public short wParamH;
        }
        public enum INPUTTYPE
        {
            Mouse,
            Keyboard,
            Hardware
        }
        [Flags]
        public enum MOUSEFLAG
        {
            MOVE = 1,
            LEFTDOWN = 2,
            LEFTUP = 4,
            RIGHTDOWN = 8,
            RIGHTUP = 16,
            MIDDLEDOWN = 32,
            MIDDLEUP = 64,
            XDOWN = 128,
            XUP = 256,
            VIRTUALDESK = 1024,
            WHEEL = 2048,
            ABSOLUTE = 32768
        }
        [Flags]
        public enum KEYBOARDFLAG
        {
            EXTENDEDKEY = 1,
            KEYUP = 2,
            UNICODE = 4,
            SCANCODE = 8
        }
        [DllImport("user32.dll", SetLastError = true)]
        public static extern int SendInput(int cInputs, ref Mouse.INPUT pInputs, int cbSize);
        [DllImport("User32")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, IntPtr dwExtraInfo);
        public static void LeftDown()
        {
            Mouse.mouse_event(2, 0, 0, 0, IntPtr.Zero);
        }
        public static void LeftUp()
        {
            Mouse.mouse_event(4, 0, 0, 0, IntPtr.Zero);
        }
        public static void DragTo(int sor_X, int sor_Y, int des_X, int des_Y)
        {
            System.Windows.Forms.Cursor.Position = new System.Drawing.Point(sor_X, sor_Y);
            Mouse.LeftDown();
            Thread.Sleep(200);
            System.Windows.Forms.Cursor.Position =  new System.Drawing.Point(des_X, des_Y);
            Mouse.LeftUp();
        }
        public static void LeftClick(int sor_X, int sor_Y)
        {
            System.Windows.Forms.Cursor.Position = new System.Drawing.Point(sor_X, sor_Y);
            Mouse.LeftDown();
            Thread.Sleep(20);
            Mouse.LeftUp();
        }
        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, UIntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hwnd, uint wMsg, int wParam, int lParam);
        [DllImport("user32.dll")]
        private static extern IntPtr PostMessage(IntPtr hWnd, uint Msg, UIntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern int PostMessage(IntPtr hwnd, uint wMsg, int wParam, int lParam);
        public static int MakeLParam(int LoWord, int HiWord)
        {
            return HiWord << 16 | (LoWord & 65535);
        }
        public static void tosleftdown(IntPtr hwnd, int sor_X, int sor_Y)
        {
            uint wMsg = 513u;
            int lParam = Mouse.MakeLParam(sor_X, sor_Y);
            Mouse.PostMessage(hwnd, wMsg, 0, lParam);
            Thread.Sleep(50);
        }
        public static void toslefup(IntPtr hwnd, int sor_X, int sor_Y)
        {
            uint wMsg = 514u;
            int lParam = Mouse.MakeLParam(sor_X, sor_Y);
            Mouse.PostMessage(hwnd, wMsg, 0, lParam);
            Thread.Sleep(20);
        }
        public static void tosdrag(IntPtr hwnd, int sor_X, int sor_Y, int time)
        {
            uint wMsg = 512u;
            int lParam = Mouse.MakeLParam(sor_X, sor_Y);
            Mouse.PostMessage(hwnd, wMsg, 514, lParam);
            Thread.Sleep(time);
        }
    }



}
