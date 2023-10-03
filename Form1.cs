using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Timers;
using System.Threading;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Automation;
using ActUtlTypeLib;



namespace communicationapplication
{
    public partial class Form1 : Form
    {

        SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\user\Desktop\testing\2communicationapplication\Database1.mdf;Integrated Security=True");
        
        public Form1()
        {
            InitializeComponent();

           // button4.Click += new EventHandler(button4_Click);
           // Button15.Click += new EventHandler(button15_Click);
        }
        ActUtlType plc = new ActUtlType() ;


        //connect to plc
        int x;
        private void button1_Click(object sender, EventArgs e)
        {
           
            plc.ActLogicalStationNumber = 5;
             x= this.plc.Open();
            if (x == 0)
            {

                commentbox.BackColor = Color.Green;
                commentbox.Text = "接続しました";
            }

            else
            {
                commentbox.BackColor = Color.Yellow;
                commentbox.Text = "接続を解除してください ";

            }


        }

        //disconnect to plc
        private void disconnect_Click(object sender, EventArgs e)
        {
            if (plc.ActLogicalStationNumber == 5)
            {
                plc.Close();
                commentbox.BackColor = Color.Yellow;
                commentbox.Text = "切断しました";
                plc.ActLogicalStationNumber = 0;
            }
            else
            {
                commentbox.Text = "エラー : 接続解除に失敗しました";
                commentbox.BackColor = Color.Yellow;

            }


        }

        //read button
        private void read_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Length > 0)
            {
                int read_result;
                plc.GetDevice(textBox1.Text, out read_result);
                textBox2.Text = read_result.ToString();
            }

            else
            {
                commentbox.Text = "入力内容のエラー";
                commentbox.BackColor = Color.Yellow;
            }

        }

        //write button
        private void write_Click(object sender, EventArgs e)
        {
            if (textBox2.Text.Length > 0)
            {
                plc.SetDevice(textBox1.Text, Convert.ToInt16(textBox2.Text));
            }
            else
            {
                commentbox.Text = "何も書くことはありません";
                commentbox.BackColor = Color.Yellow;
            }


        }
        //application open
        private void button2_Click(object sender, EventArgs e)
        {
            
            Process.Start(@"D:\T&T software\PC\Inspector.exe");
            {
                Thread.Sleep(100);
                commentbox.Text = "アプリの公開";
                commentbox.BackColor = Color.Green;
                //Thread.Sleep(100);
                //commentbox.Text = "";
                //commentbox.BackColor = Color.Silver;
            }
          

        }

        //application close
        private void button3_Click(object sender, EventArgs e)
        {
            foreach (var process in Process.GetProcessesByName("Inspector"))
            {
                process.Kill();

            }
        }
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll")]
        static public extern bool GetWindowRect(IntPtr hWnd, out Rectangle lpRect);

        [DllImport("user32.dll")]
        //private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int wMsg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern int SendNotifyMessage(IntPtr hWnd, int wMsg, IntPtr wParam, IntPtr lParam);
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        private static extern int SetForegroundWindow(IntPtr points);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern void SwitchToThisWindow(IntPtr hWnd, bool turnOn);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool SendMessageCallback(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam, SendAsyncProc lpCallBack, UIntPtr dwData);
        public delegate void SendAsyncProc(IntPtr hWnd, uint uMsg, UIntPtr dwData, IntPtr lResult);

        WaitHandle waiter = new EventWaitHandle(false, EventResetMode.ManualReset);

        const int BM_CLICK = 0x00F5;
        const int WM_LBUTTONDOWN = 0x0201;
        const int WM_LBUTTONUP = 0x0202;
        const int WM_USER = 0x400;
        const int WM_CLOSE = 0x10;

        public const int WM_SYSCOMMAND = 0x0112;
        public const int SC_CLOSE = 0xF060;

        /// <returns></returns>


        //savebutton
        private void button6_Click(object sender, EventArgs e)
        {
            commentbox.Text = "";
            commentbox.BackColor = Color.Silver;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            
                if (tableDataGridView.DataSource is DataTable dt)
                {
                    List<string> csvReady = new List<string>(dt.Rows.Count + 1);
                    csvReady.Add(string.Join(",", dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
                    List<string> rowData = new List<string>(dt.Columns.Count);
                    foreach (DataRow row in dt.Rows)
                    {
                        rowData.Clear();
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            object cell = row[i];
                            if (cell.GetType() == typeof(DateTime))
                            {
                                rowData.Add(((DateTime)cell).ToString("yyyy/MM/dd"));
                            }
                            else
                            {
                                rowData.Add(cell.ToString());
                            }
                        }
                        csvReady.Add(string.Join(",", rowData));
                    }
                    string csv = string.Join("\n", csvReady);



                    string folderPath = "C:\\CSV\\";
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }

                    {

                        File.WriteAllText(folderPath + DateTime.Now.ToString("yyyyMMdd_HH-mm-ss") + ".csv", csv, Encoding.UTF8);


                        commentbox.Text = "保存しました";
                        commentbox.BackColor = Color.Green;

                    }
                }
           
        }
    
    

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.R))
            {
                Button9.PerformClick();
                return true;
            }
            if (keyData == (Keys.Control | Keys.S))
            {
                Button7.PerformClick();
                return true;
            }
            if (keyData == (Keys.Control | Keys.C))
            {
                Button1.PerformClick();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        private void Form1_Load(object sender, EventArgs e)
        {

            DisplayData();

        }

        //reset timer 
        private void button8_Click(object sender, EventArgs e)
        {
            //if (tableDataGridView.SelectedRows.Count > 0)
            //{
            //    DialogResult dialogResult = MessageBox.Show("削除してよろしいですか Yes/No", "Delete", MessageBoxButtons.YesNo);

            //    if (dialogResult == DialogResult.Yes)
            //    {
            //        DataTable ds = new DataTable();
            //        using (SqlCommand cmd = con.CreateCommand())
            //        {

            //            con.Open();
            //            cmd.CommandText = "DELETE FROM[Table] WHERE Id=" + tableDataGridView.SelectedRows[0].Cells[0].Value.ToString() + "";
            //            cmd.CommandType = CommandType.Text;
            //            cmd.ExecuteNonQuery();
            //            con.Close();
            //            DisplayData();
            //            tableDataGridView.Update();
            //            tableDataGridView.Refresh();
            //        }

            //    }
            //}
            //else
            //{
            //    MessageBox.Show("データ行を選択してください");
            //}
            if (plc.ActLogicalStationNumber==5)
            { 
              if (commentbox.BackColor != Color.Yellow && commentbox.BackColor != Color.Red)
             {
                {
                    plc.SetDevice("M8020", Convert.ToInt16(1));


                }
                Thread.Sleep(100);
                {
                    plc.SetDevice("M8020", Convert.ToInt16(0));
                        plc.SetDevice("M8000", Convert.ToInt16(1));
                    }

                commentbox.Text = "PLCタイマーリセットしました";
                commentbox.BackColor = Color.Green;
                    plc.SetDevice("M8000", Convert.ToInt16(0));
                }
            else
            {
                error();
            }

        }
            else
            {
                commentbox.Text = "接続を確認してください。";
                commentbox.BackColor = Color.Yellow;
            }
        }



        private void tableBindingNavigatorSaveItem_Click_3(object sender, EventArgs e)
        {
            this.Validate();
            this.tableBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.data1DataSet);

        }


        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
            //}
        }
        private void ClearData()
        {
            //Id = 0;
        }

        //Read the new data
        //データを読みます
        private void button9_Click(object sender, EventArgs e)
        {
            if (plc.ActLogicalStationNumber == 5)
            {
                if (commentbox.BackColor!=Color.Yellow && commentbox.BackColor != Color.Red)
                {


                    bindingNavigatorAddNewItem.PerformClick();
                    DateTime d = new DateTime();
                    d = DateTime.Now;
                    int count;
                    plc.GetDevice("D224", out count);
                    int upvalue;
                    plc.GetDevice("D228", out upvalue);
                    int  upvalue1;
                    plc.GetDevice("D229",out upvalue1);
                    int downvalue;
                   plc.GetDevice("D227",out downvalue);
                    int downvalue1;
                    plc.GetDevice("D226", out downvalue1);
                    int tool;
                    plc.GetDevice("D225", out tool);
                    int year;
                    plc.GetDevice("D220", out year);
                    int month;
                    plc.GetDevice("D221", out  month);
                    var mo = month / 100;
                    var dy = month % 100;
                    //string s =$"{ mo}/{ dy}";
                    //int mont = year  &0xFF;
                    //int da = (year >> 8) & 0xFF;
                    //int day;
                    //plc.GetDevice("D222", out day);
                    //var dy = day % 100;
                    //string s = $"{ mo}/{ dy}";
                    int hour;
                    plc.GetDevice("D222", out hour);
                    var min = hour / 100;
                    var hrs = hour % 100;
                    //int minute;
                    // plc.GetDevice("D224", out minute);
                    int second;
                    plc.GetDevice("D223", out second);
                    var sec = second % 100;

                    double  value1 = Convert.ToDouble(upvalue);
                    double  value2 = Convert.ToDouble(upvalue1);
                    double  value3 = Convert.ToDouble(downvalue);
                    double  value4 = Convert.ToDouble(downvalue1);
                  double  x1 = value2 + value1;
                  double  y1 = value4 + value3;
                    SqlCommand cmd;

                    cmd = new SqlCommand("insert into [Table](年月日,時間分秒,打点カウント,ツールチェンジャー,上側,下側) values(@年月日,@時間分秒,@打点カウント,@ツールチェンジャー,@上側,@下側)", con);
                    con.Open();
                    {
                       // cmd.Parameters.AddWithValue("@年月日", year.ToString() + "/" + month.ToString() + "/" + month.ToString());
                        cmd.Parameters.AddWithValue("@年月日", d.ToString("yyyy/MM/dd"));
                        //cmd.Parameters.AddWithValue("@時間分秒", hrs.ToString() + ":" + min.ToString() + ":" + sec.ToString());
                      cmd.Parameters.AddWithValue("@時間分秒", d.ToString("HH:mm:ss"));
                        cmd.Parameters.AddWithValue("@打点カウント", count.ToString());
                        cmd.Parameters.AddWithValue("@ツールチェンジャー", tool.ToString());
                        cmd.Parameters.AddWithValue("@上側", x1.ToString());
                        cmd.Parameters.AddWithValue("@下側", y1.ToString());
                        cmd.ExecuteNonQuery();
                        con.Close();
                        DisplayData();
                        ClearData();
                    }
                    {
                        {
                            plc.SetDevice("M8010", Convert.ToInt16(1));
                        }
                        Thread.Sleep(50);
                        {
                            plc.SetDevice("M8010", Convert.ToInt16(0));

                        }

                    }
                    commentbox.Text = "読込完了フラグを送信しました";
                    commentbox.BackColor = Color.Green;

                }
                else
                {
                    error();
                }
            }
            else
            {
                commentbox.Text = "接続を確認してください。";
                commentbox.BackColor = Color.Yellow;
            }

        }

        private void dgZavod_RowsAdded_1(object sender, DataGridViewRowsAddedEventArgs e)
        {
            tableDataGridView.FirstDisplayedScrollingRowIndex = tableDataGridView.Rows[tableDataGridView.Rows.Count - 1].Index;
        }
        private void DisplayData()
        {
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter adp = new SqlDataAdapter("SELECT * FROM [Table] ", con);
            adp.Fill(dt);
            tableDataGridView.DataSource = dt;
            con.Close();


        }
        private void error()
        {
            commentbox.Text = "エラーを確認してください";
            commentbox.BackColor = Color.Red;

        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
        }

        //delete data from database
        //データベースからデータ消去
        private void button10_Click(object sender, EventArgs e)
        {
            if (tableDataGridView.Rows.Count > 0)
            {
                DialogResult dialogResult = MessageBox.Show("削除してよろしいですか Yes/No", "Delete", MessageBoxButtons.YesNo);

                if (dialogResult == DialogResult.Yes)
                {
                    DataTable dt = new DataTable();
                    using (SqlCommand cmd = con.CreateCommand())
                    {
                        cmd.CommandText = "TRUNCATE TABLE [Table]";
                        

                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                        DisplayData();

                        tableDataGridView.Refresh();
                    }
                    commentbox.Text = "データテーブルリセットしました";
                    commentbox.BackColor = Color.Green;
                }
               
            }
            else
            {
                commentbox.Text = "削除するデータがない";
                commentbox.BackColor = Color.Yellow;
           
            }

        }

        private void button13_Click(object sender, EventArgs e)
        {

            IntPtr maindHwnd = FindWindow(null, "Inspector");
            IntPtr maindHwnd1 = FindWindow(null, "Error");
            if (maindHwnd != IntPtr.Zero)
            {
                //IntPtr main1 = FindWindow(null, "Error");
                IntPtr panel = FindWindowEx(maindHwnd, IntPtr.Zero, "MDIClient", null);
                IntPtr panel1 = FindWindowEx(panel, IntPtr.Zero, "TAveForm", null);
                IntPtr panel2 = FindWindowEx(panel1, IntPtr.Zero, "TPanel", "Panel5");
                IntPtr panel3 = FindWindowEx(panel2, IntPtr.Zero, "TPanel", null);
                IntPtr childHwnd = FindWindowEx(panel3, IntPtr.Zero, "TBitBtn", "Save");


                if (childHwnd != IntPtr.Zero)
                {

                    SendMessage(childHwnd, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                    SetForegroundWindow(maindHwnd1);
                }
                Thread.Sleep(200);
                if (maindHwnd1 != IntPtr.Zero)
                {

                    IntPtr childHwnd1 = FindWindowEx(maindHwnd1, IntPtr.Zero, "Button", "Ok");   // Get the handle of the button
                    if (childHwnd1 != IntPtr.Zero)
                    {
                        SendMessage(childHwnd, BM_CLICK, IntPtr.Zero, IntPtr.Zero);     // Send a message from the button

                    }
                }



                else
                {
                    commentbox.BackColor = Color.Yellow;
                    commentbox.Text = "エラープログラムを再起動してください";
                }
            }
            else
            {
                commentbox.BackColor = Color.Yellow;
                commentbox.Text = "ファイルはまだ開いていません";
            }

        }
        //messagebox close
        private void button4_Click(object sender, EventArgs e)
        {

            IntPtr hWnd = FindWindow(null, "Error");

            if (hWnd != IntPtr.Zero)
            {

                IntPtr childHwnd = FindWindowEx(hWnd, IntPtr.Zero, "Button", "Ok");   // Get the handle of the button
                if (childHwnd != IntPtr.Zero)
                {
                    SendMessage(childHwnd, BM_CLICK, IntPtr.Zero, IntPtr.Zero);     // Send a message from the button

                }
                else
                {
                    commentbox.BackColor = Color.Yellow;
                    commentbox.Text = "エラー";
                }
            }
            else
            {
                commentbox.BackColor = Color.Yellow;
                commentbox.Text = "エラーメッセージボックスありません";
                //}

            }



        }

        private void button5_Click(object sender, EventArgs e)
        {

            IntPtr maindHwnd = FindWindow(null, "Connect");
            if (maindHwnd != IntPtr.Zero)
            {
                IntPtr panel = FindWindowEx(maindHwnd, IntPtr.Zero, "TPanel", null);
                IntPtr childHwnd = FindWindowEx(panel, IntPtr.Zero, null, "Link");
                if (childHwnd != IntPtr.Zero)
                {
                    SendMessage(childHwnd, BM_CLICK, IntPtr.Zero, IntPtr.Zero);     // Send a message from the button
                }
                else
                {
                    commentbox.BackColor = Color.Yellow;
                    commentbox.Text = "エラー";
                }
            }
            else
            {
                commentbox.BackColor = Color.Yellow;
                commentbox.Text = "エラーメッセージボックスありません";
            }

        }

        private void button11_Click(object sender, EventArgs e)
        {
            IntPtr maindHwnd = FindWindow(null, "Connect");
            if (maindHwnd != IntPtr.Zero)
            {
                IntPtr panel = FindWindowEx(maindHwnd, IntPtr.Zero, "TPanel", null);
                IntPtr childHwnd = FindWindowEx(panel, IntPtr.Zero, null, "Cancel");
                if (childHwnd != IntPtr.Zero)
                {
                    SendMessage(childHwnd, BM_CLICK, IntPtr.Zero, IntPtr.Zero);     // Send a message from the button
                }
                else
                {
                    commentbox.BackColor = Color.Yellow;
                    commentbox.Text = "エラー";
                }
            }
            else
            {
                commentbox.BackColor = Color.Yellow;
                commentbox.Text = "エラーメッセージボックスありません";
            }


        }

        private void button12_Click(object sender, EventArgs e)
        {
            IntPtr maindHwnd = FindWindow(null, "Inspector");
            if (maindHwnd != IntPtr.Zero)
            {
                IntPtr panel = FindWindowEx(maindHwnd, IntPtr.Zero, "TPanel", null);
                IntPtr childHwnd = FindWindowEx(panel, IntPtr.Zero, "TBitBtn", "Connect");
               
                if (childHwnd != IntPtr.Zero)
                {
                    PostMessage(childHwnd, BM_CLICK, 0, 0);     // Send a message from the button
                   // commentbox.Text = "エラー";
                }
                else
                {
                    commentbox.BackColor = Color.Yellow;
                    commentbox.Text = "エラー";
                }
            }
            else
            {
                commentbox.BackColor = Color.Yellow;
                commentbox.Text = "エラーメッセージボックスありません";
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            IntPtr main1 = FindWindow(null, "form1");
            IntPtr maindHwnd = FindWindow(null, "Inspector");
            if (maindHwnd != IntPtr.Zero)
            {
                IntPtr panel = FindWindowEx(maindHwnd, IntPtr.Zero, "MDIClient", null);
                IntPtr panel1 = FindWindowEx(panel, IntPtr.Zero, "TAveForm", null);
                IntPtr panel2 = FindWindowEx(panel1, IntPtr.Zero, "TPanel", "Panel5");
                IntPtr panel3 = FindWindowEx(panel2, IntPtr.Zero, "TPanel", null);
                IntPtr childHwnd = FindWindowEx(panel3, IntPtr.Zero, "TBitBtn", "Open");
                if (childHwnd != IntPtr.Zero)
                {

                    PostMessage(childHwnd, BM_CLICK, 0, 0);
                    

                    if (MessageBox.Show("There is no data!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.Yes)
                    {
                        this.Close();
                        MessageBox.Show("completed");
                    }
                }

            }


        }

        private void button15_Click(object sender, EventArgs e)
        {
            String textBox = DateTime.Now.ToString("yyyyMMdd_HH-mm-ss");
            IntPtr maindHwnd = FindWindow(null, "Decide file name.");
            IntPtr hWnd = FindWindow(null, "Error");
            if (maindHwnd != IntPtr.Zero)
            {
                {
                    IntPtr panel = FindWindowEx(maindHwnd, IntPtr.Zero, "ComboBoxEx32", null);
                    IntPtr panel1 = FindWindowEx(panel, IntPtr.Zero, "ComboBox", null);
                    IntPtr panel2 = FindWindowEx(panel1, IntPtr.Zero, "Edit", null);
                    if (panel2 != IntPtr.Zero)
                    {

                        SendKeys.Send(textBox);
                        SetForegroundWindow(panel2);
                    }

                }
                Thread.Sleep(200);
                IntPtr maindHwnd1 = FindWindow(null, "Decide file name.");
                if (maindHwnd1 != IntPtr.Zero)
                {
                    IntPtr childHwnd1 = FindWindowEx(maindHwnd1, IntPtr.Zero, "Button", "&Save");
                    if (childHwnd1 != IntPtr.Zero)
                    {
                        SendMessage(childHwnd1, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                        commentbox.Text = "アプリのデータ保存しました";
                    }
                }
            }
            if (hWnd != IntPtr.Zero)
            {
                IntPtr childHwnd = FindWindowEx(hWnd, IntPtr.Zero, "Button", "Ok");   // Get the handle of the button
                if (childHwnd != IntPtr.Zero)
                {
                    SendMessage(childHwnd, BM_CLICK, IntPtr.Zero, IntPtr.Zero);     // Send a message from the button

                }
                else
                {
                    commentbox.BackColor = Color.Yellow;
                    commentbox.Text = "エラー";
                }
            }

         

        }

        private void button16_Click(object sender, EventArgs e)
        {



            IntPtr maindHwnd = FindWindow(null, "Decide file name.");
            if (maindHwnd != IntPtr.Zero)
            {
                IntPtr childHwnd = FindWindowEx(maindHwnd, IntPtr.Zero, "Button", "&Save");
                if (childHwnd != IntPtr.Zero)
                {
                    SendMessage(childHwnd, BM_CLICK, IntPtr.Zero, IntPtr.Zero);     // Send a message from the button
                }
                else
                {
                    commentbox.BackColor = Color.Yellow;
                    commentbox.Text = "エラー";
                }
            }
            else
            {
                commentbox.BackColor = Color.Yellow;
                commentbox.Text = "エラーメッセージボックスありません";
            }
        }


        private void saveBtn_Click(object sender, EventArgs e)
        {
            
                IntPtr maindHwnd = FindWindow(null, "Inspector");
                if (maindHwnd != IntPtr.Zero)
                {
                    //IntPtr main1 = FindWindow(null, "Error");
                    IntPtr panel = FindWindowEx(maindHwnd, IntPtr.Zero, "MDIClient", null);
                    IntPtr panel1 = FindWindowEx(panel, IntPtr.Zero, "TAveForm", null);
                    IntPtr panel2 = FindWindowEx(panel1, IntPtr.Zero, "TPanel", "Panel5");
                    IntPtr panel3 = FindWindowEx(panel2, IntPtr.Zero, "TPanel", null);
                    IntPtr childHwnd = FindWindowEx(panel3, IntPtr.Zero, "TBitBtn", "Save");


                    if (childHwnd != IntPtr.Zero)
                    {

                        PostMessage(childHwnd, BM_CLICK, 0, 0);
                    
                    commentbox.Text = "保存ボタンクリックしました";
                    commentbox.BackColor = Color.Green;
               　　 }
                    else
                    {
                        commentbox.BackColor = Color.Yellow;
                        commentbox.Text = "ファイルはまだ開いていません";
                    }

                }
            else
            {
                
                commentbox.Text = "エラー：アプリを確認してください";
                commentbox.BackColor = Color.Yellow;
            }
 
        }
      
    }
}
