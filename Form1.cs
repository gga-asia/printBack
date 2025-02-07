using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PrintKernel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Reflection.Emit;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using Label = System.Windows.Forms.Label;
using Point = System.Drawing.Point;
using MailMessage = System.Net.Mail.MailMessage;
using Task = System.Threading.Tasks.Task;
using System.Net.Http;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using System.Diagnostics;
using System.Configuration;
using System.Collections;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        DBA DBA = new DBA();
        EmailSender EmailSender = new EmailSender();
        private ListBox listBox1;
        private ListBox listBoxErr;

        private Button controlButton;

        private WordBase1 word1 = new WordBase1();
        private WordBase2 word2 = new WordBase2();
        private WordBase3 word3 = new WordBase3();
        private WordBase4 word4 = new WordBase4();
        private WordBase5 word5 = new WordBase5();
        private WordBase6 word6 = new WordBase6();
        private WordBase7 word7 = new WordBase7();
        private WordBase8 word8 = new WordBase8();
        private WordBase9 word9 = new WordBase9();

        private static System.Timers.Timer myTimer1;
        private static System.Timers.Timer myTimer2;
        private static System.Timers.Timer myTimer3;
        private static System.Timers.Timer myTimer4;
        private static System.Timers.Timer myTimer5;
        private static System.Timers.Timer myTimer6;
        private static System.Timers.Timer myTimer7;
        private static System.Timers.Timer myTimer8;
        private static System.Timers.Timer myTimer9;
        private string watermarkText = "正式區";
        Task[] tasks = new Task[9];
        private CancellationTokenSource[] cancellationTokens = new CancellationTokenSource[9];
        private bool isTask1Running = false; // Flag to track task status
        private System.Timers.Timer clearTimer;
        private CancellationTokenSource[] stopCheckCancellations = new CancellationTokenSource[9]; // 管理每個計時器的取消源
        private bool[] isTimerRunning = new bool[9]; // 用來記錄每個計時器的運行狀態
        private System.Timers.Timer deleteDocTimer;
        private Panel[] timerStatusLights = new Panel[9];  // 用來顯示每個 Timer 狀態的燈
        private Label[] timerLabel = new Label[9];

        public Form1()
        {
            InitializeComponent();
            InitializeUI();//UI
            InitializeWatermark();//浮水印
            SetDailyDeleteTimer();//清除暫存及資料夾的WORD檔(每天晚上10點)
            InitializeClearTimer();//每小時清一次ListBox1
        }
        private void InitializeWatermark()
        {
            label1.ForeColor = Color.Gray;
            label1.Text = watermarkText;
            listBoxErr.Items.Add("事件訊息:");

        }
        private void InitializeUI()
        {
            // 初始化控制按钮
            listBoxErr = new ListBox
            {
                Dock = DockStyle.Top,
                Height = 150
            };
            this.Controls.Add(listBoxErr);

            listBox1 = new ListBox
            {
                Location = new Point(10, listBoxErr.Bottom + 5),
                Size = new Size(620, 300)      // 設置大小
            };
            this.Controls.Add(listBox1);

            controlButton = new Button
            {
                Text = "啟動",
                Dock = DockStyle.Bottom
            };
            controlButton.Click += ControlButton_Click;
            this.Controls.Add(controlButton);

            //TESTButton = new Button
            //{
            //    Text = "TEST",
            //    Location = new Point( 10, listBox1.Bottom + 60),
            //    Size = new Size(90, 30)
            //};
            //TESTButton.Click += TESTButton_Click;
            //this.Controls.Add(TESTButton);

            // 初始化Panel+Label
            for (int i = 0; i < 9; i++)
            {
                timerStatusLights[i] = new Panel
                {
                    Size = new Size(60, 20),  // 圓點大小
                    Location = new Point(10 + (i * 70), listBox1.Bottom + 5),  // 控件位置
                    BackColor = Color.Red  // 初始設置為紅色，表示未啟動
                };
                // 初始化 Label
                timerLabel[i] = new Label
                {
                    //Text = $"Timer {i + 1}",  // 設置初始文字
                    TextAlign = ContentAlignment.MiddleCenter,  // 文字居中對齊
                    Dock = DockStyle.Fill,  // 使 Label 填滿 Panel
                    ForeColor = Color.Black  // 設置文字顏色，確保與背景顏色對比清晰
                };

                // 將 Label 添加到 Panel 中
                timerStatusLights[i].Controls.Add(timerLabel[i]);

                // 將 Panel 添加到 Form 中
                this.Controls.Add(timerStatusLights[i]);
            }
        }

        public void InitializeClearTimer()
        {
            // Create a timer with a one-hour interval (3600000 milliseconds = 1 hour)
            clearTimer = new System.Timers.Timer(3600000);
            clearTimer.Elapsed += ClearListBox;  // Attach the Elapsed event to clear the ListBox
            clearTimer.AutoReset = true;         // Keep resetting the timer
            clearTimer.Start();                  // Start the timer
        }
        private void ClearListBox(object sender, System.Timers.ElapsedEventArgs e)
        {
            // Since the Timer runs in a separate thread, we need to use Invoke to update the UI safely
            if (listBox1.InvokeRequired)
            {
                listBox1.Invoke(new Action(() =>
                {
                    listBox1.Items.Clear();

                }));
            }
            else
            {
                listBox1.Items.Clear();

            }
        }
        private void TESTButton_Click(object sender, EventArgs e)
        {
            // Stop the current task
            if (cancellationTokens[0] != null)
            {
                cancellationTokens[0].Cancel();
                tasks[0]?.Wait();  // Wait for the task to cancel
            }
        }
        private void ControlButton_Click(object sender, EventArgs e)
        {
            if (controlButton.Text == "啟動")
            {
                controlButton.Text = "重新啟動 (針對已停止Timer)";
            }

            // 檢查所有 9 個計時器是否已經在運行
            bool areAllTimersRunning = true;
            for (int i = 0; i < 9; i++)
            {
                if (!isTimerRunning[i])
                {
                    areAllTimersRunning = false;
                    break;
                }
            }

            // 如果所有計時器都在運行，則無需重新啟動
            if (areAllTimersRunning)
            {
                this.Invoke(new Action(() =>
                {
                    listBoxErr.Items.Add("所有計時器已在運行，無需重新啟動 " + DateTime.Now.ToString());
                    listBoxErr.TopIndex = listBoxErr.Items.Count - 1;
                }));
                return;
            }

            // 停止所有當前的任務並重新啟動那些未運行的計時器
            for (int i = 0; i < 9; i++)
            {
                if (!isTimerRunning[i])  // 只有未運行的計時器才需要重啟
                {
                    // 停止當前的任務
                    if (cancellationTokens[i] != null)
                    {
                        cancellationTokens[i].Cancel();
                        tasks[i]?.Wait();  // 等待任務取消完成
                    }

                    // 創建新的 CancellationTokenSource 用於該計時器
                    cancellationTokens[i] = new CancellationTokenSource();

                    // 重啟該計時器的 Task
                    switch (i)
                    {
                        case 0:
                            tasks[0] = System.Threading.Tasks.Task.Run(() => Timer1_work(cancellationTokens[0].Token));
                            break;
                        case 1:
                            tasks[1] = System.Threading.Tasks.Task.Run(() => Timer2_work(cancellationTokens[1].Token));
                            break;
                        case 2:
                            tasks[2] = System.Threading.Tasks.Task.Run(() => Timer3_work(cancellationTokens[2].Token));
                            break;
                        case 3:
                            tasks[3] = System.Threading.Tasks.Task.Run(() => Timer4_work(cancellationTokens[3].Token));
                            break;
                        case 4:
                            tasks[4] = System.Threading.Tasks.Task.Run(() => Timer5_work(cancellationTokens[4].Token));
                            break;
                        case 5:
                            tasks[5] = System.Threading.Tasks.Task.Run(() => Timer6_work(cancellationTokens[5].Token));
                            break;
                        case 6:
                            tasks[6] = System.Threading.Tasks.Task.Run(() => Timer7_work(cancellationTokens[6].Token));
                            break;
                        case 7:
                            tasks[7] = System.Threading.Tasks.Task.Run(() => Timer8_work(cancellationTokens[7].Token));
                            break;
                        case 8:
                            tasks[8] = System.Threading.Tasks.Task.Run(() => Timer9_work(cancellationTokens[8].Token));
                            break;
                        default:
                            break;
                    }
                    isTimerRunning[i] = true;
                }
            }
        }
        private async Task CheckIfTimerRestarts(int timerIndex, CancellationToken token, string product)
        {
            try
            {
                // 等待10秒，看計時器是否在此期間重新啟動
                await Task.Delay(10000, token);

                // 如果10秒內未取消，表示計時器未重新啟動，發送郵件
                if (!isTimerRunning[timerIndex])
                {
                    this.Invoke((MethodInvoker)delegate
                    {
                        listBoxErr.Items.Add($"計時器 {timerIndex} 停止超過10秒");
                        listBoxErr.Items.Add("發送郵件: " + DateTime.Now.ToString());
                        listBoxErr.TopIndex = listBoxErr.Items.Count - 1; // 滾動到最後一行
                    });

                    // 調用發送郵件函數                    
                    EmailSender.SendEmailAsync("HankChang@GGA.ASIA", "HankChang@GGA.ASIA", "a0918136928@gmail.com", product + "停止超過10秒", "中斷時間" + DateTime.Now.AddSeconds(-10).ToString());

                    // 防止多次重複啟動計時器
                    isTimerRunning[timerIndex] = false;
                }

                //重新啟動
                ControlButton_Click(controlButton, EventArgs.Empty);

            }
            catch (TaskCanceledException)
            {
                // 如果計時器在10秒內重新啟動，這裡捕獲取消操作
            }
        }
        private async void Timer1_work(CancellationToken token)
        {
            // 開始時將計時器設置為運行狀態
            isTimerRunning[0] = true;

            // 燈設置為綠色，表示計時器正在運行
            this.Invoke((MethodInvoker)delegate
            {
                timerStatusLights[0].BackColor = Color.Lime;
                timerLabel[0].Text = "LIS";
            });

            myTimer1 = new System.Timers.Timer();
            //Elapsed代表,time設定的時間到之後要執行的方法
            myTimer1.Elapsed += new ElapsedEventHandler(reportTime1);
            myTimer1.Interval = 1 * 1000;
            myTimer1.Start();

            //isEmailSent[0] = false;

            try
            {
                // Wait for the task to be canceled
                await System.Threading.Tasks.Task.Delay(Timeout.Infinite, token);
            }
            catch (TaskCanceledException)
            {
                // Handle cancellation
            }

            // If cancellation is requested, stop the timer           
            myTimer1.Stop();
            myTimer1.Dispose();

            // 將計時器狀態設置為停止
            isTimerRunning[0] = false;

            // 燈設置為紅色，表示計時器停止
            this.Invoke((MethodInvoker)delegate
            {
                timerStatusLights[0].BackColor = Color.Red;
                timerLabel[0].Text = "LIS";
            });

            // 開始檢查10秒內是否重新啟動
            stopCheckCancellations[0] = new CancellationTokenSource();
            await CheckIfTimerRestarts(0, stopCheckCancellations[0].Token, "LIS");
        }
        private async void Timer2_work(CancellationToken token)
        {
            // 開始時將計時器設置為運行狀態
            isTimerRunning[1] = true;

            // 燈設置為綠色，表示計時器正在運行
            this.Invoke((MethodInvoker)delegate
            {
                timerStatusLights[1].BackColor = Color.Lime;
                timerLabel[1].Text = "SNP";
            });

            myTimer2 = new System.Timers.Timer();
            //Elapsed代表,time設定的時間到之後要執行的方法
            myTimer2.Elapsed += new ElapsedEventHandler(reportTime2);
            myTimer2.Interval = 1 * 1050;
            myTimer2.Start();

            try
            {
                // Wait for the task to be canceled
                await System.Threading.Tasks.Task.Delay(Timeout.Infinite, token);
            }
            catch (TaskCanceledException)
            {
                // Handle cancellation
            }

            // Stop the timer when cancellation is requested
            myTimer2.Stop();
            myTimer2.Dispose();

            // 將計時器狀態設置為停止
            isTimerRunning[1] = false;

            // 燈設置為紅色，表示計時器停止
            this.Invoke((MethodInvoker)delegate
            {
                timerStatusLights[1].BackColor = Color.Red;
                timerLabel[1].Text = "SNP";
            });

            // 開始檢查10秒內是否重新啟動
            stopCheckCancellations[1] = new CancellationTokenSource();
            await CheckIfTimerRestarts(1, stopCheckCancellations[1].Token, "SNP");
        }
        private async void Timer3_work(CancellationToken token)
        {
            // 開始時將計時器設置為運行狀態
            isTimerRunning[2] = true;

            // 燈設置為綠色，表示計時器正在運行
            this.Invoke((MethodInvoker)delegate
            {
                timerStatusLights[2].BackColor = Color.Lime;
                timerLabel[2].Text = "NOT SNP";
            });

            myTimer3 = new System.Timers.Timer();
            //Elapsed代表,time設定的時間到之後要執行的方法
            myTimer3.Elapsed += new ElapsedEventHandler(reportTime3);
            myTimer3.Interval = 1 * 1100;
            myTimer3.Start();

            try
            {
                // Wait for the task to be canceled
                await System.Threading.Tasks.Task.Delay(Timeout.Infinite, token);
            }
            catch (TaskCanceledException)
            {
                // Handle cancellation
            }
            // Stop the timer when cancellation is requested
            myTimer3.Stop();
            myTimer3.Dispose();

            // 將計時器狀態設置為停止
            isTimerRunning[2] = false;

            // 燈設置為紅色，表示計時器停止
            this.Invoke((MethodInvoker)delegate
            {
                timerStatusLights[2].BackColor = Color.Red;
                timerLabel[2].Text = "NOT SNP";
            });

            // 開始檢查10秒內是否重新啟動
            stopCheckCancellations[2] = new CancellationTokenSource();
            await CheckIfTimerRestarts(2, stopCheckCancellations[2].Token, "NotSNP");
        }
        private async void Timer4_work(CancellationToken token)
        {
            // 開始時將計時器設置為運行狀態
            isTimerRunning[3] = true;

            // 燈設置為綠色，表示計時器正在運行
            this.Invoke((MethodInvoker)delegate
            {
                timerStatusLights[3].BackColor = Color.Lime;
                timerLabel[3].Text = "FFD 0-4";
            });

            myTimer4 = new System.Timers.Timer();
            //Elapsed代表,time設定的時間到之後要執行的方法
            myTimer4.Elapsed += new ElapsedEventHandler(reportTime4);
            myTimer4.Interval = 1 * 1150;
            myTimer4.Start();

            try
            {
                // Wait for the task to be canceled
                await System.Threading.Tasks.Task.Delay(Timeout.Infinite, token);
            }
            catch (TaskCanceledException)
            {
                // Handle cancellation
            }

            // Stop the timer when cancellation is requested
            myTimer4.Stop();
            myTimer4.Dispose();

            // 將計時器狀態設置為停止
            isTimerRunning[3] = false;

            // 燈設置為紅色，表示計時器停止
            this.Invoke((MethodInvoker)delegate
            {
                timerStatusLights[3].BackColor = Color.Red;
                timerLabel[3].Text = "FFD 0-4";
            });

            // 開始檢查10秒內是否重新啟動
            stopCheckCancellations[3] = new CancellationTokenSource();
            await CheckIfTimerRestarts(3, stopCheckCancellations[3].Token, "FFD 0-4");
        }
        private async void Timer5_work(CancellationToken token)
        {
            // 開始時將計時器設置為運行狀態
            isTimerRunning[4] = true;

            // 燈設置為綠色，表示計時器正在運行
            this.Invoke((MethodInvoker)delegate
            {
                timerStatusLights[4].BackColor = Color.Lime;
                timerLabel[4].Text = "FFD 5-9";
            });

            myTimer5 = new System.Timers.Timer();
            //Elapsed代表,time設定的時間到之後要執行的方法           
            myTimer5.Elapsed += new ElapsedEventHandler(reportTime5);
            myTimer5.Interval = 1 * 1200;
            myTimer5.Start();
            try
            {
                // Wait for the task to be canceled
                await System.Threading.Tasks.Task.Delay(Timeout.Infinite, token);
            }
            catch (TaskCanceledException)
            {
                // Handle cancellation
            }

            // Stop the timer when cancellation is requested
            myTimer5.Stop();
            myTimer5.Dispose();

            // 將計時器狀態設置為停止
            isTimerRunning[4] = false;

            // 燈設置為紅色，表示計時器停止
            this.Invoke((MethodInvoker)delegate
            {
                timerStatusLights[4].BackColor = Color.Red;
                timerLabel[4].Text = "FFD 5-9";
            });

            // 開始檢查10秒內是否重新啟動
            stopCheckCancellations[4] = new CancellationTokenSource();
            await CheckIfTimerRestarts(4, stopCheckCancellations[4].Token, "FFD 5-9");
        }
        private async void Timer6_work(CancellationToken token)
        {
            // 開始時將計時器設置為運行狀態
            isTimerRunning[5] = true;

            // 燈設置為綠色，表示計時器正在運行
            this.Invoke((MethodInvoker)delegate
            {
                timerStatusLights[5].BackColor = Color.Lime;
                timerLabel[5].Text = "FTS 0-4";
            });

            myTimer6 = new System.Timers.Timer();
            //Elapsed代表,time設定的時間到之後要執行的方法
            myTimer6.Elapsed += new ElapsedEventHandler(reportTime6);
            myTimer6.Interval = 1 * 1250;
            myTimer6.Start();

            try
            {
                // Wait for the task to be canceled
                await System.Threading.Tasks.Task.Delay(Timeout.Infinite, token);
            }
            catch (TaskCanceledException)
            {
                // Handle cancellation
            }

            // Stop the timer when cancellation is requested
            myTimer6.Stop();
            myTimer6.Dispose();

            // 將計時器狀態設置為停止
            isTimerRunning[5] = false;

            // 燈設置為紅色，表示計時器停止
            this.Invoke((MethodInvoker)delegate
            {
                timerStatusLights[5].BackColor = Color.Red;
                timerLabel[5].Text = "FTS 0-4";
            });

            // 開始檢查10秒內是否重新啟動
            stopCheckCancellations[5] = new CancellationTokenSource();
            await CheckIfTimerRestarts(5, stopCheckCancellations[5].Token, "FTS 0-4");
        }
        private async void Timer7_work(CancellationToken token)
        {
            // 開始時將計時器設置為運行狀態
            isTimerRunning[6] = true;

            // 燈設置為綠色，表示計時器正在運行
            this.Invoke((MethodInvoker)delegate
            {
                timerStatusLights[6].BackColor = Color.Lime;
                timerLabel[6].Text = "FTS 5-9";
            });

            myTimer7 = new System.Timers.Timer();
            //Elapsed代表,time設定的時間到之後要執行的方法
            myTimer7.Elapsed += new ElapsedEventHandler(reportTime7);
            myTimer7.Interval = 1 * 1300;
            myTimer7.Start();

            try
            {
                // Wait for the task to be canceled
                await System.Threading.Tasks.Task.Delay(Timeout.Infinite, token);
            }
            catch (TaskCanceledException)
            {
                // Handle cancellation
            }
            // Stop the timer when cancellation is requested
            myTimer7.Stop();
            myTimer7.Dispose();

            // 將計時器狀態設置為停止
            isTimerRunning[6] = false;

            // 燈設置為紅色，表示計時器停止
            this.Invoke((MethodInvoker)delegate
            {
                timerStatusLights[6].BackColor = Color.Red;
                timerLabel[6].Text = "FTS 5-9";
            });

            // 開始檢查10秒內是否重新啟動
            stopCheckCancellations[6] = new CancellationTokenSource();
            await CheckIfTimerRestarts(6, stopCheckCancellations[6].Token, "FTS 5-9");
        }
        private async void Timer8_work(CancellationToken token)
        {
            // 開始時將計時器設置為運行狀態
            isTimerRunning[7] = true;

            // 燈設置為綠色，表示計時器正在運行
            this.Invoke((MethodInvoker)delegate
            {
                timerStatusLights[7].BackColor = Color.Lime;
                timerLabel[7].Text = "SMA 0-4";
            });

            myTimer8 = new System.Timers.Timer();
            //Elapsed代表,time設定的時間到之後要執行的方法
            myTimer8.Elapsed += new ElapsedEventHandler(reportTime8);
            myTimer8.Interval = 1 * 1350;
            myTimer8.Start();

            try
            {
                // Wait for the task to be canceled
                await System.Threading.Tasks.Task.Delay(Timeout.Infinite, token);
            }
            catch (TaskCanceledException)
            {
                // Handle cancellation
            }

            // Stop the timer when cancellation is requested
            myTimer8.Stop();
            myTimer8.Dispose();

            // 將計時器狀態設置為停止
            isTimerRunning[7] = false;

            // 燈設置為紅色，表示計時器停止
            this.Invoke((MethodInvoker)delegate
            {
                timerStatusLights[7].BackColor = Color.Red;
                timerLabel[7].Text = "SMA 0-4";
            });

            // 開始檢查10秒內是否重新啟動
            stopCheckCancellations[7] = new CancellationTokenSource();
            await CheckIfTimerRestarts(7, stopCheckCancellations[7].Token, "SMA 0-4");
        }
        private async void Timer9_work(CancellationToken token)
        {
            // 開始時將計時器設置為運行狀態
            isTimerRunning[8] = true;

            // 燈設置為綠色，表示計時器正在運行
            this.Invoke((MethodInvoker)delegate
            {
                timerStatusLights[8].BackColor = Color.Lime;
                timerLabel[8].Text = "SMA 5-9";
            });

            myTimer9 = new System.Timers.Timer();
            //Elapsed代表,time設定的時間到之後要執行的方法           
            myTimer9.Elapsed += new ElapsedEventHandler(reportTime9);
            myTimer9.Interval = 1 * 1400;
            myTimer9.Start();

            try
            {
                // Wait for the task to be canceled
                await System.Threading.Tasks.Task.Delay(Timeout.Infinite, token);
            }
            catch (TaskCanceledException)
            {
                // Handle cancellation
            }

            // Stop the timer when cancellation is requested
            myTimer9.Stop();
            myTimer9.Dispose();

            // 將計時器狀態設置為停止
            isTimerRunning[8] = false;

            // 燈設置為紅色，表示計時器停止
            this.Invoke((MethodInvoker)delegate
            {
                timerStatusLights[8].BackColor = Color.Red;
                timerLabel[8].Text = "SMA 5-9";
            });

            // 開始檢查10秒內是否重新啟動
            stopCheckCancellations[8] = new CancellationTokenSource();
            await CheckIfTimerRestarts(8, stopCheckCancellations[8].Token, "SMA 5-9");
        }

        /// <summary>
        /// LIS
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void reportTime1(object sender, ElapsedEventArgs e)
        {
            string logMessage = "LIS: " + DateTime.Now.ToString();
            this.Invoke(new Action(() =>
            {
                listBox1.Items.Add(logMessage);
                listBox1.TopIndex = listBox1.Items.Count - 1;
            }));

            TimerParameters parameters = new TimerParameters();
            parameters = GetTimerParametersFromDatabase(1);
            string P_Text = "";
            //listBox.Items.Add("timer1_start:" + DateTime.Now.ToString());
            if (parameters.Product.IndexOf(',') > -1)
            {
                string[] P_list = parameters.Product.Split(',');
                foreach (var P_item in P_list)
                {
                    if (P_Text == "")
                    {
                        if (parameters.P_Type == "1")
                        {
                            P_Text = " FilePath NOT like '%" + P_item + "%' ";
                        }
                        else
                        {
                            P_Text = " FilePath like '%" + P_item + "%' ";
                        }
                    }
                    else
                    {
                        if (parameters.P_Type == "1")
                        {
                            P_Text = P_Text + " AND  FilePath NOT like '%" + P_item + "%' ";
                        }
                        else
                        {
                            P_Text = P_Text + " OR FilePath like '%" + P_item + "%' ";
                        }
                    }
                }
            }
            else
            {
                if (parameters.P_Type == "1")
                {
                    P_Text = " FilePath NOT like '%" + parameters.Product + "%' ";
                }
                else
                {
                    P_Text = " FilePath like '%" + parameters.Product + "%' ";
                }
            }
            string Num_Text = "";
            if (parameters.PrintNum.IndexOf(',') > -1)
            {
                string[] Num_list = parameters.PrintNum.Split(',');
                foreach (var Num_item in Num_list)
                {
                    if (Num_Text == "")
                    {
                        Num_Text = " FilePath LIKE '%" + Num_item + ".doc' OR FilePath LIKE '%" + Num_item + ".pdf' ";
                    }
                    else
                    {
                        Num_Text = Num_Text + " OR FilePath LIKE '%" + Num_item + ".doc' OR FilePath LIKE '%" + Num_item + ".pdf' ";
                    }
                }
            }

            //檢查Queue Table是否有待套印筆數           
            //Console.WriteLine(string.Format("{0} cycle checking.", DateTime.Now));
            string sql = "select top (1) * from LIS_QUEUE_MASTER_" + parameters.C_System + " nolock ";
            //sql += " where QueueID = '202410171804405812' ";
            sql += " where Flag = 'N' ";
            sql += " and TemplateFileName not like '%.rpt%' ";
            sql += " And (" + P_Text + ") ";
            if (Num_Text != "")
            {
                sql += " AND ( " + Num_Text + " ) ";
            }
            sql += " order by CreateDateTime ";

            System.Data.DataTable dt_master = DBA.GetDataTable(sql);
            if (this.DBA.LastError != "")
            {
                //Console.WriteLine(this.DBA.LastError);
                string ErrMessage = "LIS: " + this.DBA.LastError;
                this.Invoke(new Action(() =>
                {
                    listBoxErr.Items.Add(ErrMessage);
                }));
            }

            if (dt_master != null && dt_master.Rows.Count > 0)
            {

                //待套印筆數
                foreach (DataRow dr_m in dt_master.Rows)
                {
                    String QueueID = dr_m["QueueID"].ToString().Trim();
                    String TargetFilePath = dr_m["FilePath"].ToString().Trim();
                    try
                    {
                        SetStatus("Start", QueueID, parameters.C_System, "word1");
                        DateTime startTime = DateTime.Now;    //轉換起始時間
                        List<string> sPDFFileNameList = new List<string>();
                        //listBox.Items.Add(string.Format("{0} 開始轉換", startTime));
                        //Word範本路徑
                        string TemplateFilePath = dr_m["TemplateFileName"].ToString().Trim();

                        //檔案路徑
                        string TargetFileName = dr_m["FilePath"].ToString().Trim();

                        //JSON資料還原成DataTable
                        System.Data.DataTable Dt = JsonConvert.DeserializeObject<System.Data.DataTable>(dr_m["ReportData"].ToString().Trim());

                        //檔案格式(PDF 或 WORD)
                        string FileType = dr_m["FileType"].ToString().Trim();

                        //是否為暫存檔案
                        string IsTempFile = dr_m["IsTempFile"].ToString().Trim();
                        myTimer1.Stop();
                        //開始匯出
                        this.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add("從範本複製一份到新檔案(LIS)");
                        }));
                        this.word1.Export(TemplateFilePath, TargetFileName, Dt, FileType, IsTempFile);


                        if (String.IsNullOrEmpty(this.word1.LastError))
                        {
                            SetStatus("Finish", QueueID, parameters.C_System, "word1");
                        }
                        else
                        {
                            SetStatus("Error", QueueID, parameters.C_System, "word1");
                        }
                        this.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add("檔案轉換成功(LIS)");
                        }));
                        DateTime endtime = DateTime.Now;    //轉換結束時間
                        TimeSpan duration = new TimeSpan(endtime.Ticks - startTime.Ticks);  //計算一筆花多少時間

                        //listBox.Items.Add(string.Format("{0} 檔名:{1} 花費{2}秒 ", endtime, TargetFilePath, duration.Seconds));
                        //listBox.Items.Add("-------------------------------------------------");
                    }
                    catch (Exception ex)
                    {
                        string ERRMessage = $"例外訊息: {ex.Message}\n堆疊追蹤: {ex.StackTrace}";
                        SetStatus("Error", QueueID, parameters.C_System, "word1");
                        InsertLog(QueueID, ERRMessage);
                        EmailSender.SendEmailAsync(ConfigurationManager.AppSettings["MailTo"], "", "", $"列印核心異常({QueueID}) {TargetFilePath}", ERRMessage);
                        DeleteFailReport(TargetFilePath, ConfigurationManager.AppSettings["LISFilePath"]);
                    }
                    finally
                    {
                        word1.Init();
                        //釋放資源 
                        word1.Close();
                    }
                }

                myTimer1.Start();
            }
        }
        /// <summary>
        /// SNP
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void reportTime2(object sender, ElapsedEventArgs e)
        {
            string logMessage = "SNP: " + DateTime.Now.ToString();
            this.Invoke(new Action(() =>
            {
                listBox1.Items.Add(logMessage);
                listBox1.TopIndex = listBox1.Items.Count - 1;
            }));
            TimerParameters parameters1 = new TimerParameters();
            parameters1 = GetTimerParametersFromDatabase(2);
            string P_Text = "";
            //listBox.Items.Add("timer2_start:" + DateTime.Now.ToString());
            if (parameters1.Product.IndexOf(',') > -1)
            {
                string[] P_list = parameters1.Product.Split(',');
                foreach (var P_item in P_list)
                {
                    if (P_Text == "")
                    {
                        if (parameters1.P_Type == "1")
                        {
                            P_Text = " FilePath NOT like '%" + P_item + "%' ";
                        }
                        else
                        {
                            P_Text = " FilePath like '%" + P_item + "%' ";
                        }
                    }
                    else
                    {
                        if (parameters1.P_Type == "1")
                        {
                            P_Text = P_Text + " AND  FilePath NOT like '%" + P_item + "%' ";
                        }
                        else
                        {
                            P_Text = P_Text + " OR FilePath like '%" + P_item + "%' ";
                        }
                    }
                }
            }
            else
            {
                if (parameters1.P_Type == "1")
                {
                    P_Text = " FilePath NOT like '%" + parameters1.Product + "%' ";
                }
                else
                {
                    P_Text = " FilePath like '%" + parameters1.Product + "%' ";
                }
            }
            string Num_Text = "";
            if (parameters1.PrintNum.IndexOf(',') > -1)
            {
                string[] Num_list = parameters1.PrintNum.Split(',');
                foreach (var Num_item in Num_list)
                {
                    if (Num_Text == "")
                    {
                        Num_Text = " FilePath LIKE '%" + Num_item + ".doc' OR FilePath LIKE '%" + Num_item + ".pdf' ";
                    }
                    else
                    {
                        Num_Text = Num_Text + " OR FilePath LIKE '%" + Num_item + ".doc' OR FilePath LIKE '%" + Num_item + ".pdf' ";
                    }
                }
            }

            //檢查Queue Table是否有待套印筆數           
            //Console.WriteLine(string.Format("{0} cycle checking.", DateTime.Now));
            string sql = "select top (1) * from LIS_QUEUE_MASTER_" + parameters1.C_System + " nolock ";
            //sql += " where QueueID = '20240412142013931' ";
            sql += " where Flag = 'N' ";
            sql += " and TemplateFileName not like '%.rpt%' ";
            sql += " And (" + P_Text + ") ";
            if (Num_Text != "")
            {
                sql += " AND ( " + Num_Text + " ) ";
            }
            sql += " order by CreateDateTime ";

            System.Data.DataTable dt_master = DBA.GetDataTable(sql);
            if (this.DBA.LastError != "")
            {
                //Console.WriteLine(this.DBA.LastError);
                string ErrMessage = "SNP: " + this.DBA.LastError;
                this.Invoke(new Action(() =>
                {
                    listBoxErr.Items.Add(ErrMessage);
                }));
            }

            if (dt_master != null && dt_master.Rows.Count > 0)
            {

                //待套印筆數
                foreach (DataRow dr_m in dt_master.Rows)
                {
                    String QueueID = dr_m["QueueID"].ToString().Trim();
                    String TargetFilePath = dr_m["FilePath"].ToString().Trim();
                    try
                    {
                        SetStatus("Start", QueueID, parameters1.C_System, "word2");
                        DateTime startTime = DateTime.Now;    //轉換起始時間
                        List<string> sPDFFileNameList = new List<string>();
                        //listBox.Items.Add(string.Format("{0} 開始轉換", startTime));
                        //Word範本路徑
                        string TemplateFilePath = dr_m["TemplateFileName"].ToString().Trim();

                        //檔案路徑
                        string TargetFileName = dr_m["FilePath"].ToString().Trim();

                        //JSON資料還原成DataTable
                        System.Data.DataTable Dt = JsonConvert.DeserializeObject<System.Data.DataTable>(dr_m["ReportData"].ToString().Trim());

                        //檔案格式(PDF 或 WORD)
                        string FileType = dr_m["FileType"].ToString().Trim();

                        //是否為暫存檔案
                        string IsTempFile = dr_m["IsTempFile"].ToString().Trim();
                        myTimer2.Stop();
                        //開始匯出
                        this.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add("從範本複製一份到新檔案(SNP)");
                        }));
                        this.word2.Export(TemplateFilePath, TargetFileName, Dt, FileType, IsTempFile);



                        if (String.IsNullOrEmpty(this.word2.LastError))
                        {
                            SetStatus("Finish", QueueID, parameters1.C_System, "word2");
                        }
                        else
                        {
                            SetStatus("Error", QueueID, parameters1.C_System, "word2");
                        }
                        this.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add("檔案轉換成功(SNP)");
                        }));
                        DateTime endtime = DateTime.Now;    //轉換結束時間
                        TimeSpan duration = new TimeSpan(endtime.Ticks - startTime.Ticks);  //計算一筆花多少時間

                        //listBox.Items.Add(string.Format("{0} 檔名:{1} 花費{2}秒 ", endtime, TargetFilePath, duration.Seconds));
                        //listBox.Items.Add("-------------------------------------------------");
                    }
                    catch (Exception ex)
                    {
                        string ERRMessage = $"例外訊息: {ex.Message}\n堆疊追蹤: {ex.StackTrace}";
                        SetStatus("Error", QueueID, parameters1.C_System, "word2");
                        InsertLog(QueueID, ERRMessage);
                        EmailSender.SendEmailAsync(ConfigurationManager.AppSettings["MailTo"], "", "", $"列印核心異常({QueueID}) {TargetFilePath}", ERRMessage);
                        DeleteFailReport(TargetFilePath, ConfigurationManager.AppSettings["GGALISFilePath"]);
                    }
                    finally
                    {
                        word2.Init();
                        //釋放資源 
                        word2.Close();
                    }
                    myTimer2.Start();
                }

            }

        }

        /// <summary>
        /// NOT SNP
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void reportTime3(object sender, ElapsedEventArgs e)
        {
            string logMessage = "NOT SNP: " + DateTime.Now.ToString();
            this.Invoke(new Action(() =>
            {
                listBox1.Items.Add(logMessage);
                listBox1.TopIndex = listBox1.Items.Count - 1;
            }));
            TimerParameters parameters2 = new TimerParameters();
            parameters2 = GetTimerParametersFromDatabase(3);
            string P_Text = "";
            //listBox.Items.Add("timer3_start:" + DateTime.Now.ToString());
            if (parameters2.Product.IndexOf(',') > -1)
            {
                string[] P_list = parameters2.Product.Split(',');
                foreach (var P_item in P_list)
                {
                    if (P_Text == "")
                    {
                        if (parameters2.P_Type == "1")
                        {
                            P_Text = " FilePath NOT like '%" + P_item + "%' ";
                        }
                        else
                        {
                            P_Text = " FilePath like '%" + P_item + "%' ";
                        }
                    }
                    else
                    {
                        if (parameters2.P_Type == "1")
                        {
                            P_Text = P_Text + " AND  FilePath NOT like '%" + P_item + "%' ";
                        }
                        else
                        {
                            P_Text = P_Text + " OR FilePath like '%" + P_item + "%' ";
                        }
                    }
                }
            }
            else
            {
                if (parameters2.P_Type == "1")
                {
                    P_Text = " FilePath NOT like '%" + parameters2.Product + "%' ";
                }
                else
                {
                    P_Text = " FilePath like '%" + parameters2.Product + "%' ";
                }
            }
            string Num_Text = "";
            if (parameters2.PrintNum.IndexOf(',') > -1)
            {
                string[] Num_list = parameters2.PrintNum.Split(',');
                foreach (var Num_item in Num_list)
                {
                    if (Num_Text == "")
                    {
                        Num_Text = " FilePath LIKE '%" + Num_item + ".doc' OR FilePath LIKE '%" + Num_item + ".pdf' ";
                    }
                    else
                    {
                        Num_Text = Num_Text + " OR FilePath LIKE '%" + Num_item + ".doc' OR FilePath LIKE '%" + Num_item + ".pdf' ";
                    }
                }
            }

            //檢查Queue Table是否有待套印筆數           
            //Console.WriteLine(string.Format("{0} cycle checking.", DateTime.Now));
            string sql = "select TOP (1) * from LIS_QUEUE_MASTER_" + parameters2.C_System + " nolock ";
            //sql += " where QueueID = '20241028101234526' ";
            sql += " where Flag = 'N' ";
            sql += " and TemplateFileName not like '%.rpt%' ";
            sql += " And (" + P_Text + ") ";
            if (Num_Text != "")
            {
                sql += " AND ( " + Num_Text + " ) ";
            }
            sql += " order by CreateDateTime ";

            System.Data.DataTable dt_master = DBA.GetDataTable(sql);
            if (this.DBA.LastError != "")
            {
                //Console.WriteLine(this.DBA.LastError);            
                string ErrMessage = "NOT SNP: " + this.DBA.LastError;
                this.Invoke(new Action(() =>
                {
                    listBoxErr.Items.Add(ErrMessage);
                }));
            }

            if (dt_master != null && dt_master.Rows.Count > 0)
            {

                //待套印筆數
                foreach (DataRow dr_m in dt_master.Rows)
                {
                    String QueueID = dr_m["QueueID"].ToString().Trim();
                    String TargetFilePath = dr_m["FilePath"].ToString().Trim();
                    try
                    {
                        SetStatus("Start", QueueID, parameters2.C_System, "word3");
                        DateTime startTime = DateTime.Now;    //轉換起始時間
                        List<string> sPDFFileNameList = new List<string>();
                        //listBox.Items.Add(string.Format("{0} 開始轉換", startTime));
                        //Word範本路徑
                        string TemplateFilePath = dr_m["TemplateFileName"].ToString().Trim();

                        //檔案路徑
                        string TargetFileName = dr_m["FilePath"].ToString().Trim();

                        //JSON資料還原成DataTable
                        System.Data.DataTable Dt = JsonConvert.DeserializeObject<System.Data.DataTable>(dr_m["ReportData"].ToString().Trim());

                        //檔案格式(PDF 或 WORD)
                        string FileType = dr_m["FileType"].ToString().Trim();

                        //是否為暫存檔案
                        string IsTempFile = dr_m["IsTempFile"].ToString().Trim();
                        myTimer3.Stop();
                        //開始匯出
                        this.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add("從範本複製一份到新檔案(NOT SNP)");
                        }));
                        this.word3.Export(TemplateFilePath, TargetFileName, Dt, FileType, IsTempFile);


                        if (String.IsNullOrEmpty(this.word3.LastError))
                        {
                            SetStatus("Finish", QueueID, parameters2.C_System, "word3");
                        }
                        else
                        {
                            SetStatus("Error", QueueID, parameters2.C_System, "word3");
                        }
                        this.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add("檔案轉換成功(NOT SNP)");
                        }));
                        DateTime endtime = DateTime.Now;    //轉換結束時間
                        TimeSpan duration = new TimeSpan(endtime.Ticks - startTime.Ticks);  //計算一筆花多少時間

                        //listBox.Items.Add(string.Format("{0} 檔名:{1} 花費{2}秒 ", endtime, TargetFilePath, duration.Seconds));
                        //listBox.Items.Add("-------------------------------------------------");
                    }
                    catch (Exception ex)
                    {
                        string ERRMessage = $"例外訊息: {ex.Message}\n堆疊追蹤: {ex.StackTrace}";
                        SetStatus("Error", QueueID, parameters2.C_System, "word3");
                        InsertLog(QueueID, ERRMessage);
                        EmailSender.SendEmailAsync(ConfigurationManager.AppSettings["MailTo"], "", "", $"列印核心異常({QueueID}) {TargetFilePath}", ERRMessage);
                        DeleteFailReport(TargetFilePath, ConfigurationManager.AppSettings["GGALISFilePath"]);
                    }
                    finally
                    {
                        word3.Init();
                        //釋放資源 
                        word3.Close();
                    }
                    myTimer3.Start();
                }

            }

        }
        /// <summary>
        /// FFD 0-4
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void reportTime4(object sender, ElapsedEventArgs e)
        {
            string logMessage = "FFD 0-4: " + DateTime.Now.ToString();
            this.Invoke(new Action(() =>
            {
                listBox1.Items.Add(logMessage);
                listBox1.TopIndex = listBox1.Items.Count - 1;
            }));
            TimerParameters parameters3 = new TimerParameters();
            parameters3 = GetTimerParametersFromDatabase(4);
            string P_Text = "";
            //listBox.Items.Add("timer4_start:" + DateTime.Now.ToString());
            if (parameters3.Product.IndexOf(',') > -1)
            {
                string[] P_list = parameters3.Product.Split(',');
                foreach (var P_item in P_list)
                {
                    if (P_Text == "")
                    {
                        if (parameters3.P_Type == "1")
                        {
                            P_Text = " FilePath NOT like '%" + P_item + "%' ";
                        }
                        else
                        {
                            P_Text = " FilePath like '%" + P_item + "%' ";
                        }
                    }
                    else
                    {
                        if (parameters3.P_Type == "1")
                        {
                            P_Text = P_Text + " AND  FilePath NOT like '%" + P_item + "%' ";
                        }
                        else
                        {
                            P_Text = P_Text + " OR FilePath like '%" + P_item + "%' ";
                        }
                    }
                }
            }
            else
            {
                if (parameters3.P_Type == "1")
                {
                    P_Text = " FilePath NOT like '%" + parameters3.Product + "%' ";
                }
                else
                {
                    P_Text = " FilePath like '%" + parameters3.Product + "%' ";
                }
            }
            string Num_Text = "";
            if (parameters3.PrintNum.IndexOf(',') > -1)
            {
                string[] Num_list = parameters3.PrintNum.Split(',');
                foreach (var Num_item in Num_list)
                {
                    if (Num_Text == "")
                    {
                        Num_Text = " FilePath LIKE '%" + Num_item + ".doc' OR FilePath LIKE '%" + Num_item + ".pdf' ";
                    }
                    else
                    {
                        Num_Text = Num_Text + " OR FilePath LIKE '%" + Num_item + ".doc' OR FilePath LIKE '%" + Num_item + ".pdf' ";
                    }
                }
            }

            //檢查Queue Table是否有待套印筆數           
            //Console.WriteLine(string.Format("{0} cycle checking.", DateTime.Now));
            string sql = "select TOP (1) * from LIS_QUEUE_MASTER_" + parameters3.C_System + " nolock ";
            //sql += " where QueueID = '202410241733401352' ";
            sql += " where Flag = 'N' ";
            sql += " and TemplateFileName not like '%.rpt%' ";
            sql += " And (" + P_Text + ") ";
            if (Num_Text != "")
            {
                sql += " AND ( " + Num_Text + " ) ";
            }
            sql += " order by CreateDateTime ";

            System.Data.DataTable dt_master = DBA.GetDataTable(sql);
            if (this.DBA.LastError != "")
            {
                //Console.WriteLine(this.DBA.LastError);
                string ErrMessage = "FFD 0-4: " + this.DBA.LastError;
                this.Invoke(new Action(() =>
                {
                    listBoxErr.Items.Add(ErrMessage);
                }));
            }

            if (dt_master != null && dt_master.Rows.Count > 0)
            {

                //待套印筆數
                foreach (DataRow dr_m in dt_master.Rows)
                {
                    String QueueID = dr_m["QueueID"].ToString().Trim();
                    String TargetFilePath = dr_m["FilePath"].ToString().Trim();
                    try
                    {
                        SetStatus("Start", QueueID, parameters3.C_System, "word4");
                        DateTime startTime = DateTime.Now;    //轉換起始時間
                        List<string> sPDFFileNameList = new List<string>();
                        //listBox.Items.Add(string.Format("{0} 開始轉換", startTime));
                        //Word範本路徑
                        string TemplateFilePath = dr_m["TemplateFileName"].ToString().Trim();

                        //檔案路徑
                        string TargetFileName = dr_m["FilePath"].ToString().Trim();

                        //JSON資料還原成DataTable
                        System.Data.DataTable Dt = JsonConvert.DeserializeObject<System.Data.DataTable>(dr_m["ReportData"].ToString().Trim());

                        //檔案格式(PDF 或 WORD)
                        string FileType = dr_m["FileType"].ToString().Trim();

                        //是否為暫存檔案
                        string IsTempFile = dr_m["IsTempFile"].ToString().Trim();
                        myTimer4.Stop();
                        //開始匯出
                        this.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add("從範本複製一份到新檔案(FFD 0-4)");
                        }));
                        this.word4.Export(TemplateFilePath, TargetFileName, Dt, FileType, IsTempFile);


                        if (String.IsNullOrEmpty(this.word4.LastError))
                        {
                            SetStatus("Finish", QueueID, parameters3.C_System, "word4");
                        }
                        else
                        {
                            SetStatus("Error", QueueID, parameters3.C_System, "word4");
                        }

                        DateTime endtime = DateTime.Now;    //轉換結束時間
                        TimeSpan duration = new TimeSpan(endtime.Ticks - startTime.Ticks);  //計算一筆花多少時間
                        this.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add("檔案轉換成功(FFD 0-4)");
                        }));
                        //listBox.Items.Add(string.Format("{0} 檔名:{1} 花費{2}秒 ", endtime, TargetFilePath, duration.Seconds));
                        //listBox.Items.Add("-------------------------------------------------");
                    }
                    catch (Exception ex)
                    {
                        string ERRMessage = $"例外訊息: {ex.Message}\n堆疊追蹤: {ex.StackTrace}";
                        SetStatus("Error", QueueID, parameters3.C_System, "word4");
                        InsertLog(QueueID, ERRMessage);
                        EmailSender.SendEmailAsync(ConfigurationManager.AppSettings["MailTo"], "", "", $"列印核心異常({QueueID}) {TargetFilePath}", ERRMessage);
                        DeleteFailReport(TargetFilePath, ConfigurationManager.AppSettings["LISFilePath"]);
                    }
                    finally
                    {
                        word4.Init();
                        //釋放資源 
                        word4.Close();
                    }
                    myTimer4.Start();
                }

            }

        }
        /// <summary>
        /// FFD 5-9
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void reportTime5(object sender, ElapsedEventArgs e)
        {
            string logMessage = "FFD 5-9: " + DateTime.Now.ToString();
            this.Invoke(new Action(() =>
            {
                listBox1.Items.Add(logMessage);
                listBox1.TopIndex = listBox1.Items.Count - 1;
            }));
            TimerParameters parameters4 = new TimerParameters();
            parameters4 = GetTimerParametersFromDatabase(5);

            string P_Text = "";
            //listBox.Items.Add("timer5_start:" + DateTime.Now.ToString());
            if (parameters4.Product.IndexOf(',') > -1)
            {
                string[] P_list = parameters4.Product.Split(',');
                foreach (var P_item in P_list)
                {
                    if (P_Text == "")
                    {
                        if (parameters4.P_Type == "1")
                        {
                            P_Text = " FilePath NOT like '%" + P_item + "%' ";
                        }
                        else
                        {
                            P_Text = " FilePath like '%" + P_item + "%' ";
                        }
                    }
                    else
                    {
                        if (parameters4.P_Type == "1")
                        {
                            P_Text = P_Text + " AND  FilePath NOT like '%" + P_item + "%' ";
                        }
                        else
                        {
                            P_Text = P_Text + " OR FilePath like '%" + P_item + "%' ";
                        }
                    }
                }
            }
            else
            {
                if (parameters4.P_Type == "1")
                {
                    P_Text = " FilePath NOT like '%" + parameters4.Product + "%' ";
                }
                else
                {
                    P_Text = " FilePath like '%" + parameters4.Product + "%' ";
                }
            }
            string Num_Text = "";
            if (parameters4.PrintNum.IndexOf(',') > -1)
            {
                string[] Num_list = parameters4.PrintNum.Split(',');
                foreach (var Num_item in Num_list)
                {
                    if (Num_Text == "")
                    {
                        Num_Text = " FilePath LIKE '%" + Num_item + ".doc' OR FilePath LIKE '%" + Num_item + ".pdf' ";
                    }
                    else
                    {
                        Num_Text = Num_Text + " OR FilePath LIKE '%" + Num_item + ".doc' OR FilePath LIKE '%" + Num_item + ".pdf' ";
                    }
                }
            }

            //檢查Queue Table是否有待套印筆數           
            //Console.WriteLine(string.Format("{0} cycle checking.", DateTime.Now));
            string sql = "select TOP (1) * from LIS_QUEUE_MASTER_" + parameters4.C_System + " nolock ";
            //sql += " where QueueID = '202410241733401039' ";
            sql += " where Flag = 'N' ";
            sql += " and TemplateFileName not like '%.rpt%' ";
            sql += " And (" + P_Text + ") ";
            if (Num_Text != "")
            {
                sql += " AND ( " + Num_Text + " ) ";
            }
            sql += " order by CreateDateTime ";

            System.Data.DataTable dt_master = DBA.GetDataTable(sql);
            if (this.DBA.LastError != "")
            {
                //Console.WriteLine(this.DBA.LastError);
                string ErrMessage = "FFD 5-9: " + this.DBA.LastError;
                this.Invoke(new Action(() =>
                {
                    listBoxErr.Items.Add(ErrMessage);
                }));
            }

            if (dt_master != null && dt_master.Rows.Count > 0)
            {

                //待套印筆數
                foreach (DataRow dr_m in dt_master.Rows)
                {
                    String QueueID = dr_m["QueueID"].ToString().Trim();
                    String TargetFilePath = dr_m["FilePath"].ToString().Trim();
                    try
                    {
                        SetStatus("Start", QueueID, parameters4.C_System, "word5");
                        DateTime startTime = DateTime.Now;    //轉換起始時間
                        List<string> sPDFFileNameList = new List<string>();
                        //listBox.Items.Add(string.Format("{0} 開始轉換", startTime));
                        //Word範本路徑
                        string TemplateFilePath = dr_m["TemplateFileName"].ToString().Trim();

                        //檔案路徑
                        string TargetFileName = dr_m["FilePath"].ToString().Trim();

                        //JSON資料還原成DataTable
                        System.Data.DataTable Dt = JsonConvert.DeserializeObject<System.Data.DataTable>(dr_m["ReportData"].ToString().Trim());

                        //檔案格式(PDF 或 WORD)
                        string FileType = dr_m["FileType"].ToString().Trim();

                        //是否為暫存檔案
                        string IsTempFile = dr_m["IsTempFile"].ToString().Trim();
                        // 暫停計時器
                        myTimer5.Stop();
                        //開始匯出
                        this.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add("從範本複製一份到新檔案(FFD 5-9)");
                        }));
                        this.word5.Export(TemplateFilePath, TargetFileName, Dt, FileType, IsTempFile);

                        if (String.IsNullOrEmpty(this.word5.LastError))
                        {
                            SetStatus("Finish", QueueID, parameters4.C_System, "word5");
                        }
                        else
                        {
                            SetStatus("Error", QueueID, parameters4.C_System, "word5");
                        }
                        this.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add("檔案轉換成功(FFD 5-9)");
                        }));
                        DateTime endtime = DateTime.Now;    //轉換結束時間
                        TimeSpan duration = new TimeSpan(endtime.Ticks - startTime.Ticks);  //計算一筆花多少時間

                        //listBox.Items.Add(string.Format("{0} 檔名:{1} 花費{2}秒 ", endtime, TargetFilePath, duration.Seconds));
                        //listBox.Items.Add("-------------------------------------------------");
                    }
                    catch (Exception ex)
                    {
                        string ERRMessage = $"例外訊息: {ex.Message}\n堆疊追蹤: {ex.StackTrace}";
                        SetStatus("Error", QueueID, parameters4.C_System, "word5");
                        InsertLog(QueueID, ERRMessage);
                        EmailSender.SendEmailAsync(ConfigurationManager.AppSettings["MailTo"], "", "", $"列印核心異常({QueueID}) {TargetFilePath}", ERRMessage);
                        DeleteFailReport(TargetFilePath, ConfigurationManager.AppSettings["LISFilePath"]);
                    }
                    finally
                    {
                        word5.Init();
                        //釋放資源 
                        word5.Close();
                    }
                    myTimer5.Start();
                }


            }
        }
        /// <summary>
        /// FTS 0-4
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void reportTime6(object sender, ElapsedEventArgs e)
        {
            string logMessage = "FTS 0-4: " + DateTime.Now.ToString();
            this.Invoke(new Action(() =>
            {
                listBox1.Items.Add(logMessage);
                listBox1.TopIndex = listBox1.Items.Count - 1;
            }));
            TimerParameters parameters4 = new TimerParameters();
            parameters4 = GetTimerParametersFromDatabase(6);

            string P_Text = "";
            //listBox.Items.Add("timer5_start:" + DateTime.Now.ToString());
            if (parameters4.Product.IndexOf(',') > -1)
            {
                string[] P_list = parameters4.Product.Split(',');
                foreach (var P_item in P_list)
                {
                    if (P_Text == "")
                    {
                        if (parameters4.P_Type == "1")
                        {
                            P_Text = " FilePath NOT like '%" + P_item + "%' ";
                        }
                        else
                        {
                            P_Text = " FilePath like '%" + P_item + "%' ";
                        }
                    }
                    else
                    {
                        if (parameters4.P_Type == "1")
                        {
                            P_Text = P_Text + " AND  FilePath NOT like '%" + P_item + "%' ";
                        }
                        else
                        {
                            P_Text = P_Text + " OR FilePath like '%" + P_item + "%' ";
                        }
                    }
                }
            }
            else
            {
                if (parameters4.P_Type == "1")
                {
                    P_Text = " FilePath NOT like '%" + parameters4.Product + "%' ";
                }
                else
                {
                    P_Text = " FilePath like '%" + parameters4.Product + "%' ";
                }
            }
            string Num_Text = "";
            if (parameters4.PrintNum.IndexOf(',') > -1)
            {
                string[] Num_list = parameters4.PrintNum.Split(',');
                foreach (var Num_item in Num_list)
                {
                    if (Num_Text == "")
                    {
                        Num_Text = " FilePath LIKE '%" + Num_item + ".doc' OR FilePath LIKE '%" + Num_item + ".pdf' ";
                    }
                    else
                    {
                        Num_Text = Num_Text + " OR FilePath LIKE '%" + Num_item + ".doc' OR FilePath LIKE '%" + Num_item + ".pdf' ";
                    }
                }
            }

            //檢查Queue Table是否有待套印筆數           
            //Console.WriteLine(string.Format("{0} cycle checking.", DateTime.Now));
            string sql = "select TOP (1) * from LIS_QUEUE_MASTER_" + parameters4.C_System + " nolock ";
            //sql += " where QueueID = '202409251034555044' ";
            sql += " where Flag = 'N' ";
            sql += " and TemplateFileName not like '%.rpt%' ";
            sql += " And (" + P_Text + ") ";
            if (Num_Text != "")
            {
                sql += " AND ( " + Num_Text + " ) ";
            }
            sql += " order by CreateDateTime ";

            System.Data.DataTable dt_master = DBA.GetDataTable(sql);
            if (this.DBA.LastError != "")
            {
                //Console.WriteLine(this.DBA.LastError);
                string ErrMessage = "FTS 0-4: " + this.DBA.LastError;
                this.Invoke(new Action(() =>
                {
                    listBoxErr.Items.Add(ErrMessage);
                }));
            }

            if (dt_master != null && dt_master.Rows.Count > 0)
            {

                //待套印筆數
                foreach (DataRow dr_m in dt_master.Rows)
                {
                    String QueueID = dr_m["QueueID"].ToString().Trim();
                    String TargetFilePath = dr_m["FilePath"].ToString().Trim();
                    try
                    {
                        SetStatus("Start", QueueID, parameters4.C_System, "word6");
                        DateTime startTime = DateTime.Now;    //轉換起始時間
                        List<string> sPDFFileNameList = new List<string>();
                        //listBox.Items.Add(string.Format("{0} 開始轉換", startTime));
                        //Word範本路徑
                        string TemplateFilePath = dr_m["TemplateFileName"].ToString().Trim();

                        //檔案路徑
                        string TargetFileName = dr_m["FilePath"].ToString().Trim();

                        //JSON資料還原成DataTable
                        System.Data.DataTable Dt = JsonConvert.DeserializeObject<System.Data.DataTable>(dr_m["ReportData"].ToString().Trim());

                        //檔案格式(PDF 或 WORD)
                        string FileType = dr_m["FileType"].ToString().Trim();

                        //是否為暫存檔案
                        string IsTempFile = dr_m["IsTempFile"].ToString().Trim();
                        // 暫停計時器
                        myTimer6.Stop();
                        //開始匯出
                        this.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add("從範本複製一份到新檔案(FTS 0-4)");
                        }));
                        this.word6.Export(TemplateFilePath, TargetFileName, Dt, FileType, IsTempFile);



                        if (String.IsNullOrEmpty(this.word6.LastError))
                        {
                            SetStatus("Finish", QueueID, parameters4.C_System, "word6");
                        }
                        else
                        {
                            SetStatus("Error", QueueID, parameters4.C_System, "word6");
                        }
                        this.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add("檔案轉換成功(FTS 0-4)");
                        }));
                        DateTime endtime = DateTime.Now;    //轉換結束時間
                        TimeSpan duration = new TimeSpan(endtime.Ticks - startTime.Ticks);  //計算一筆花多少時間

                        //listBox.Items.Add(string.Format("{0} 檔名:{1} 花費{2}秒 ", endtime, TargetFilePath, duration.Seconds));
                        //listBox.Items.Add("-------------------------------------------------");
                    }
                    catch (Exception ex)
                    {
                        string ERRMessage = $"例外訊息: {ex.Message}\n堆疊追蹤: {ex.StackTrace}";
                        SetStatus("Error", QueueID, parameters4.C_System, "word6");
                        InsertLog(QueueID, ERRMessage);
                        EmailSender.SendEmailAsync(ConfigurationManager.AppSettings["MailTo"], "", "", $"列印核心異常({QueueID}) {TargetFilePath}", ERRMessage);
                        DeleteFailReport(TargetFilePath, ConfigurationManager.AppSettings["LISFilePath"]);
                    }
                    finally
                    {
                        word6.Init();
                        //釋放資源 
                        word6.Close();
                    }
                    myTimer6.Start();
                }


            }
        }
        /// <summary>
        /// FTS 5-9
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void reportTime7(object sender, ElapsedEventArgs e)
        {
            string logMessage = "FTS 5-9: " + DateTime.Now.ToString();
            this.Invoke(new Action(() =>
            {
                listBox1.Items.Add(logMessage);
                listBox1.TopIndex = listBox1.Items.Count - 1;
            }));
            TimerParameters parameters4 = new TimerParameters();
            parameters4 = GetTimerParametersFromDatabase(7);

            string P_Text = "";
            //listBox.Items.Add("timer5_start:" + DateTime.Now.ToString());
            if (parameters4.Product.IndexOf(',') > -1)
            {
                string[] P_list = parameters4.Product.Split(',');
                foreach (var P_item in P_list)
                {
                    if (P_Text == "")
                    {
                        if (parameters4.P_Type == "1")
                        {
                            P_Text = " FilePath NOT like '%" + P_item + "%' ";
                        }
                        else
                        {
                            P_Text = " FilePath like '%" + P_item + "%' ";
                        }
                    }
                    else
                    {
                        if (parameters4.P_Type == "1")
                        {
                            P_Text = P_Text + " AND  FilePath NOT like '%" + P_item + "%' ";
                        }
                        else
                        {
                            P_Text = P_Text + " OR FilePath like '%" + P_item + "%' ";
                        }
                    }
                }
            }
            else
            {
                if (parameters4.P_Type == "1")
                {
                    P_Text = " FilePath NOT like '%" + parameters4.Product + "%' ";
                }
                else
                {
                    P_Text = " FilePath like '%" + parameters4.Product + "%' ";
                }
            }
            string Num_Text = "";
            if (parameters4.PrintNum.IndexOf(',') > -1)
            {
                string[] Num_list = parameters4.PrintNum.Split(',');
                foreach (var Num_item in Num_list)
                {
                    if (Num_Text == "")
                    {
                        Num_Text = " FilePath LIKE '%" + Num_item + ".doc' OR FilePath LIKE '%" + Num_item + ".pdf' ";
                    }
                    else
                    {
                        Num_Text = Num_Text + " OR FilePath LIKE '%" + Num_item + ".doc' OR FilePath LIKE '%" + Num_item + ".pdf' ";
                    }
                }
            }

            //檢查Queue Table是否有待套印筆數           
            //Console.WriteLine(string.Format("{0} cycle checking.", DateTime.Now));
            string sql = "select TOP (1) * from LIS_QUEUE_MASTER_" + parameters4.C_System + " nolock ";
            //sql += " where QueueID = '202409250835262658' ";
            sql += " where Flag = 'N' ";
            sql += " and TemplateFileName not like '%.rpt%' ";
            sql += " And (" + P_Text + ") ";
            if (Num_Text != "")
            {
                sql += " AND ( " + Num_Text + " ) ";
            }
            sql += " order by CreateDateTime ";

            System.Data.DataTable dt_master = DBA.GetDataTable(sql);
            if (this.DBA.LastError != "")
            {
                //Console.WriteLine(this.DBA.LastError);
                string ErrMessage = "FTS 5-9: " + this.DBA.LastError;
                this.Invoke(new Action(() =>
                {
                    listBoxErr.Items.Add(ErrMessage);
                }));
            }

            if (dt_master != null && dt_master.Rows.Count > 0)
            {

                //待套印筆數
                foreach (DataRow dr_m in dt_master.Rows)
                {
                    String QueueID = dr_m["QueueID"].ToString().Trim();
                    String TargetFilePath = dr_m["FilePath"].ToString().Trim();
                    try
                    {
                        SetStatus("Start", QueueID, parameters4.C_System, "word7");
                        DateTime startTime = DateTime.Now;    //轉換起始時間
                        List<string> sPDFFileNameList = new List<string>();
                        //listBox.Items.Add(string.Format("{0} 開始轉換", startTime));
                        //Word範本路徑
                        string TemplateFilePath = dr_m["TemplateFileName"].ToString().Trim();

                        //檔案路徑
                        string TargetFileName = dr_m["FilePath"].ToString().Trim();

                        //JSON資料還原成DataTable
                        System.Data.DataTable Dt = JsonConvert.DeserializeObject<System.Data.DataTable>(dr_m["ReportData"].ToString().Trim());

                        //檔案格式(PDF 或 WORD)
                        string FileType = dr_m["FileType"].ToString().Trim();

                        //是否為暫存檔案
                        string IsTempFile = dr_m["IsTempFile"].ToString().Trim();
                        // 暫停計時器
                        myTimer7.Stop();
                        //開始匯出
                        this.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add("從範本複製一份到新檔案(FTS 5-9)");
                        }));
                        this.word7.Export(TemplateFilePath, TargetFileName, Dt, FileType, IsTempFile);

                        if (String.IsNullOrEmpty(this.word7.LastError))
                        {
                            SetStatus("Finish", QueueID, parameters4.C_System, "word7");
                        }
                        else
                        {
                            SetStatus("Error", QueueID, parameters4.C_System, "word7");
                        }
                        this.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add("檔案轉換成功(FTS 5-9)");
                        }));
                        DateTime endtime = DateTime.Now;    //轉換結束時間
                        TimeSpan duration = new TimeSpan(endtime.Ticks - startTime.Ticks);  //計算一筆花多少時間

                        //listBox.Items.Add(string.Format("{0} 檔名:{1} 花費{2}秒 ", endtime, TargetFilePath, duration.Seconds));
                        //listBox.Items.Add("-------------------------------------------------");
                    }
                    catch (Exception ex)
                    {
                        string ERRMessage = $"例外訊息: {ex.Message}\n堆疊追蹤: {ex.StackTrace}";
                        SetStatus("Error", QueueID, parameters4.C_System, "word7");
                        InsertLog(QueueID, ERRMessage);
                        EmailSender.SendEmailAsync(ConfigurationManager.AppSettings["MailTo"], "", "", $"列印核心異常({QueueID}) {TargetFilePath}", ERRMessage);
                        DeleteFailReport(TargetFilePath, ConfigurationManager.AppSettings["LISFilePath"]);
                    }
                    finally
                    {
                        word7.Init();
                        //釋放資源 
                        word7.Close();
                    }
                    myTimer7.Start();
                }

            }
        }
        /// <summary>
        /// SMA 0-4
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void reportTime8(object sender, ElapsedEventArgs e)
        {
            string logMessage = "SMA 0-4: " + DateTime.Now.ToString();
            this.Invoke(new Action(() =>
            {
                listBox1.Items.Add(logMessage);
                listBox1.TopIndex = listBox1.Items.Count - 1;
            }));
            TimerParameters parameters4 = new TimerParameters();
            parameters4 = GetTimerParametersFromDatabase(8);

            string P_Text = "";
            //listBox.Items.Add("timer5_start:" + DateTime.Now.ToString());
            if (parameters4.Product.IndexOf(',') > -1)
            {
                string[] P_list = parameters4.Product.Split(',');
                foreach (var P_item in P_list)
                {
                    if (P_Text == "")
                    {
                        if (parameters4.P_Type == "1")
                        {
                            P_Text = " FilePath NOT like '%" + P_item + "%' ";
                        }
                        else
                        {
                            P_Text = " FilePath like '%" + P_item + "%' ";
                        }
                    }
                    else
                    {
                        if (parameters4.P_Type == "1")
                        {
                            P_Text = P_Text + " AND  FilePath NOT like '%" + P_item + "%' ";
                        }
                        else
                        {
                            P_Text = P_Text + " OR FilePath like '%" + P_item + "%' ";
                        }
                    }
                }
            }
            else
            {
                if (parameters4.P_Type == "1")
                {
                    P_Text = " FilePath NOT like '%" + parameters4.Product + "%' ";
                }
                else
                {
                    P_Text = " FilePath like '%" + parameters4.Product + "%' ";
                }
            }
            string Num_Text = "";
            if (parameters4.PrintNum.IndexOf(',') > -1)
            {
                string[] Num_list = parameters4.PrintNum.Split(',');
                foreach (var Num_item in Num_list)
                {
                    if (Num_Text == "")
                    {
                        Num_Text = " FilePath LIKE '%" + Num_item + ".doc' OR FilePath LIKE '%" + Num_item + ".pdf' ";
                    }
                    else
                    {
                        Num_Text = Num_Text + " OR FilePath LIKE '%" + Num_item + ".doc' OR FilePath LIKE '%" + Num_item + ".pdf' ";
                    }
                }
            }

            //檢查Queue Table是否有待套印筆數           
            //Console.WriteLine(string.Format("{0} cycle checking.", DateTime.Now));
            string sql = "select TOP (1) * from LIS_QUEUE_MASTER_" + parameters4.C_System + " nolock ";
            //sql += " where QueueID = '202403081111044098' ";
            sql += " where Flag = 'N' ";
            sql += " and TemplateFileName not like '%.rpt%' ";
            sql += " And (" + P_Text + ") ";
            if (Num_Text != "")
            {
                sql += " AND ( " + Num_Text + " ) ";
            }
            sql += " order by CreateDateTime ";

            System.Data.DataTable dt_master = DBA.GetDataTable(sql);
            if (this.DBA.LastError != "")
            {
                string ErrMessage = "SMA 0-4: " + this.DBA.LastError;
                this.Invoke(new Action(() =>
                {
                    listBoxErr.Items.Add(ErrMessage);
                }));
            }

            if (dt_master != null && dt_master.Rows.Count > 0)
            {

                //待套印筆數
                foreach (DataRow dr_m in dt_master.Rows)
                {
                    String QueueID = dr_m["QueueID"].ToString().Trim();
                    String TargetFilePath = dr_m["FilePath"].ToString().Trim();
                    try
                    {
                        SetStatus("Start", QueueID, parameters4.C_System, "word8");
                        DateTime startTime = DateTime.Now;    //轉換起始時間
                        List<string> sPDFFileNameList = new List<string>();
                        //listBox.Items.Add(string.Format("{0} 開始轉換", startTime));
                        //Word範本路徑
                        string TemplateFilePath = dr_m["TemplateFileName"].ToString().Trim();

                        //檔案路徑
                        string TargetFileName = dr_m["FilePath"].ToString().Trim();

                        //JSON資料還原成DataTable
                        System.Data.DataTable Dt = JsonConvert.DeserializeObject<System.Data.DataTable>(dr_m["ReportData"].ToString().Trim());

                        //檔案格式(PDF 或 WORD)
                        string FileType = dr_m["FileType"].ToString().Trim();

                        //是否為暫存檔案
                        string IsTempFile = dr_m["IsTempFile"].ToString().Trim();
                        // 暫停計時器
                        myTimer8.Stop();
                        //開始匯出
                        this.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add("從範本複製一份到新檔案(SMA 0-4)");
                        }));
                        this.word8.Export(TemplateFilePath, TargetFileName, Dt, FileType, IsTempFile);

                        if (String.IsNullOrEmpty(this.word8.LastError))
                        {
                            SetStatus("Finish", QueueID, parameters4.C_System, "word8");
                        }
                        else
                        {
                            SetStatus("Error", QueueID, parameters4.C_System, "word8");
                        }
                        this.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add("檔案轉換成功(SMA 0-4)");
                        }));
                        DateTime endtime = DateTime.Now;    //轉換結束時間
                        TimeSpan duration = new TimeSpan(endtime.Ticks - startTime.Ticks);  //計算一筆花多少時間

                        //listBox.Items.Add(string.Format("{0} 檔名:{1} 花費{2}秒 ", endtime, TargetFilePath, duration.Seconds));
                        //listBox.Items.Add("-------------------------------------------------");
                    }
                    catch (Exception ex)
                    {
                        string ERRMessage = $"例外訊息: {ex.Message}\n堆疊追蹤: {ex.StackTrace}";
                        SetStatus("Error", QueueID, parameters4.C_System, "word8");
                        InsertLog(QueueID, ERRMessage);
                        EmailSender.SendEmailAsync(ConfigurationManager.AppSettings["MailTo"], "", "", $"列印核心異常({QueueID}) {TargetFilePath}", ERRMessage);
                        DeleteFailReport(TargetFilePath, ConfigurationManager.AppSettings["LISFilePath"]);
                    }
                    finally
                    {
                        word8.Init();
                        //釋放資源 
                        word8.Close();
                    }
                    myTimer8.Start();
                }


            }
        }
        /// <summary>
        /// SMA 5-9
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void reportTime9(object sender, ElapsedEventArgs e)
        {
            string logMessage = "SMA 5-9: " + DateTime.Now.ToString();
            this.Invoke(new Action(() =>
            {
                listBox1.Items.Add(logMessage);
                listBox1.TopIndex = listBox1.Items.Count - 1;
            }));
            TimerParameters parameters4 = new TimerParameters();
            parameters4 = GetTimerParametersFromDatabase(9);

            string P_Text = "";
            //listBox.Items.Add("timer5_start:" + DateTime.Now.ToString());
            if (parameters4.Product.IndexOf(',') > -1)
            {
                string[] P_list = parameters4.Product.Split(',');
                foreach (var P_item in P_list)
                {
                    if (P_Text == "")
                    {
                        if (parameters4.P_Type == "1")
                        {
                            P_Text = " FilePath NOT like '%" + P_item + "%' ";
                        }
                        else
                        {
                            P_Text = " FilePath like '%" + P_item + "%' ";
                        }
                    }
                    else
                    {
                        if (parameters4.P_Type == "1")
                        {
                            P_Text = P_Text + " AND  FilePath NOT like '%" + P_item + "%' ";
                        }
                        else
                        {
                            P_Text = P_Text + " OR FilePath like '%" + P_item + "%' ";
                        }
                    }
                }
            }
            else
            {
                if (parameters4.P_Type == "1")
                {
                    P_Text = " FilePath NOT like '%" + parameters4.Product + "%' ";
                }
                else
                {
                    P_Text = " FilePath like '%" + parameters4.Product + "%' ";
                }
            }
            string Num_Text = "";
            if (parameters4.PrintNum.IndexOf(',') > -1)
            {
                string[] Num_list = parameters4.PrintNum.Split(',');
                foreach (var Num_item in Num_list)
                {
                    if (Num_Text == "")
                    {
                        Num_Text = " FilePath LIKE '%" + Num_item + ".doc' OR FilePath LIKE '%" + Num_item + ".pdf' ";
                    }
                    else
                    {
                        Num_Text = Num_Text + " OR FilePath LIKE '%" + Num_item + ".doc' OR FilePath LIKE '%" + Num_item + ".pdf' ";
                    }
                }
            }

            //檢查Queue Table是否有待套印筆數           
            //Console.WriteLine(string.Format("{0} cycle checking.", DateTime.Now));
            string sql = "select TOP (1) * from LIS_QUEUE_MASTER_" + parameters4.C_System + " nolock ";
            //sql += " where QueueID = '202403181057576949' ";
            sql += " where Flag = 'N' ";
            sql += " and TemplateFileName not like '%.rpt%' ";
            sql += " And (" + P_Text + ") ";
            if (Num_Text != "")
            {
                sql += " AND ( " + Num_Text + " ) ";
            }
            sql += " order by CreateDateTime ";

            System.Data.DataTable dt_master = DBA.GetDataTable(sql);
            if (this.DBA.LastError != "")
            {
                string ErrMessage = "SMA 5-9: " + this.DBA.LastError;
                this.Invoke(new Action(() =>
                {
                    listBoxErr.Items.Add(ErrMessage);
                }));
            }

            if (dt_master != null && dt_master.Rows.Count > 0)
            {

                //待套印筆數
                foreach (DataRow dr_m in dt_master.Rows)
                {
                    String QueueID = dr_m["QueueID"].ToString().Trim();
                    String TargetFilePath = dr_m["FilePath"].ToString().Trim();
                    try
                    {
                        SetStatus("Start", QueueID, parameters4.C_System, "word9");
                        DateTime startTime = DateTime.Now;    //轉換起始時間
                        List<string> sPDFFileNameList = new List<string>();
                        //listBox.Items.Add(string.Format("{0} 開始轉換", startTime));
                        //Word範本路徑
                        string TemplateFilePath = dr_m["TemplateFileName"].ToString().Trim();

                        //檔案路徑
                        string TargetFileName = dr_m["FilePath"].ToString().Trim();

                        //JSON資料還原成DataTable
                        System.Data.DataTable Dt = JsonConvert.DeserializeObject<System.Data.DataTable>(dr_m["ReportData"].ToString().Trim());

                        //檔案格式(PDF 或 WORD)
                        string FileType = dr_m["FileType"].ToString().Trim();

                        //是否為暫存檔案
                        string IsTempFile = dr_m["IsTempFile"].ToString().Trim();
                        // 暫停計時器
                        myTimer9.Stop();
                        //開始匯出
                        this.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add("從範本複製一份到新檔案(SMA 5-9)");
                        }));
                        this.word9.Export(TemplateFilePath, TargetFileName, Dt, FileType, IsTempFile);

                        if (String.IsNullOrEmpty(this.word9.LastError))
                        {
                            SetStatus("Finish", QueueID, parameters4.C_System, "word9");
                        }
                        else
                        {
                            SetStatus("Error", QueueID, parameters4.C_System, "word9");
                        }

                        this.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add("檔案轉換成功(SMA 5-9)");
                        }));
                        DateTime endtime = DateTime.Now;    //轉換結束時間
                        TimeSpan duration = new TimeSpan(endtime.Ticks - startTime.Ticks);  //計算一筆花多少時間

                        //listBox.Items.Add(string.Format("{0} 檔名:{1} 花費{2}秒 ", endtime, TargetFilePath, duration.Seconds));
                        //listBox.Items.Add("-------------------------------------------------");
                    }
                    catch (Exception ex)
                    {
                        string ERRMessage = $"例外訊息: {ex.Message}\n堆疊追蹤: {ex.StackTrace}";                        
                        SetStatus("Error", QueueID, parameters4.C_System, "word9");
                        InsertLog(QueueID, ERRMessage);
                        EmailSender.SendEmailAsync(ConfigurationManager.AppSettings["MailTo"], "", "", $"列印核心異常({QueueID}) {TargetFilePath}", ERRMessage);
                        DeleteFailReport(TargetFilePath, ConfigurationManager.AppSettings["LISFilePath"]);
                    }
                    finally
                    {
                        word9.Init();
                        //釋放資源 
                        word9.Close();
                    }
                    myTimer9.Start();
                }


            }
        }
        public void DeleteFailReport(string FilePath,string FROM)
        {
            //ConfigurationManager.AppSettings["LISFilePath"];
            //ConfigurationManager.AppSettings["GGALISFilePath"];
            if(!string.IsNullOrEmpty(FilePath) && !string.IsNullOrEmpty(FROM))
            {
                string PATH = FROM + FilePath;
                FileDelete(PATH);
            }
        }
        public void FileDelete(string Path)
        {
            if (!string.IsNullOrEmpty(Path))
            {
                try
                {
                    if (File.Exists(Path))
                    {
                        File.Delete(Path);
                    }
                }
                catch (Exception ex)
                {
                    string ERRMessage = $"例外訊息: {ex.Message}\n堆疊追蹤: {ex.StackTrace}";
                    InsertLog($"FileDelete : {Path}", ERRMessage);
                }
            }
        }
        public void InsertLog(string QueueID, string pMsg)
        {
            try
            {
                Console.WriteLine(pMsg);
                string LogPath = string.Format(@"{0}\ErrorLog", System.Windows.Forms.Application.StartupPath);
                CheckFolder(LogPath);
                string logFile = string.Format(@"{0}\{1}.txt", LogPath, DateTime.Now.ToString("yyyy-MM-dd"));
                using (StreamWriter sw = (System.IO.File.Exists(logFile)) ? System.IO.File.AppendText(logFile) : System.IO.File.CreateText(logFile))
                {
                    sw.WriteLine("{0}  QueueID:{1}  錯誤:{2}", DateTime.Now, QueueID, pMsg);
                }
            }
            catch (Exception)
            {
            }
        }
        public bool SetStatus(string pStatus, string pQueueID, string c_System, string wordapp)
        {
            //try
            //{
                List<InputPara> para = new List<InputPara>();
                para.Add(new InputPara { name = "@QueueID", value = pQueueID, dbtype = SqlDbType.VarChar });
                string UpdateSql = "";
                if (pStatus == "Start")
                {
                    //處理中註記:S
                    para.Add(new InputPara { name = "@Flag", value = "S", dbtype = SqlDbType.VarChar });
                    UpdateSql = "Update LIS_QUEUE_MASTER_" + c_System + " Set Flag = @Flag where QueueID = @QueueID";
                }
                else if (pStatus == "Finish")
                {
                    //完成註記:Y
                    para.Add(new InputPara { name = "@Flag", value = "Y", dbtype = SqlDbType.VarChar });
                    para.Add(new InputPara { name = "@Result", value = "", dbtype = SqlDbType.VarChar });
                    para.Add(new InputPara { name = "@FinishDateTime", value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), dbtype = SqlDbType.VarChar });
                    UpdateSql = "Update LIS_QUEUE_MASTER_" + c_System + " Set Flag = @Flag,Result = @Result,FinishDateTime=@FinishDateTime where QueueID = @QueueID";

                }
                else if (pStatus == "Error")
                {
                    //錯誤註記:F  錯誤記錄:Result
                    para.Add(new InputPara { name = "@Flag", value = "E", dbtype = SqlDbType.VarChar });
                    para.Add(new InputPara { name = "@FinishDateTime", value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), dbtype = SqlDbType.VarChar });
                    var wordErrors = new Dictionary<string, string>
                        {
                            { "word1", this.word1.LastError },
                            { "word2", this.word2.LastError },
                            { "word3", this.word3.LastError },
                            { "word4", this.word4.LastError },
                            { "word5", this.word5.LastError },
                            { "word6", this.word6.LastError },
                            { "word7", this.word7.LastError },
                            { "word8", this.word8.LastError },
                            { "word9", this.word9.LastError }
                         };

                    if (wordErrors.TryGetValue(wordapp, out var error))
                    {
                        if (!string.IsNullOrEmpty(error) && error.Length > 1000)
                        {
                            error = error.Substring(0, 1000);
                        }
                        para.Add(new InputPara { name = "@Result", value = error, dbtype = SqlDbType.VarChar });
                    }
                    UpdateSql = "Update LIS_QUEUE_MASTER_" + c_System + " Set Flag = @Flag,Result = @Result,FinishDateTime=@FinishDateTime where QueueID = @QueueID";
                }

                DBA.ExeCuteNonQuery(UpdateSql, inputpara: para);

            //}
            //catch (Exception)
            //{

            //    return false;
            //}

            return true;
        }
        private void CheckFolder(string pPath)
        {
            bool exists = System.IO.Directory.Exists(pPath);

            if (!exists)
                System.IO.Directory.CreateDirectory(pPath);
        }

        // 設置每天的定時刪除器
        private void SetDailyDeleteTimer()
        {
            // 計算當前時間和晚上10點的時間差
            TimeSpan timeUntil10PM = GetTimeUntilNext10PM();

            // 設置計時器，在時間差結束後觸發
            deleteDocTimer = new System.Timers.Timer(timeUntil10PM.TotalMilliseconds);
            deleteDocTimer.Elapsed += new ElapsedEventHandler(OnDeleteDocFiles);
            deleteDocTimer.AutoReset = false; // 只觸發一次，然後再重設計時器
            deleteDocTimer.Start();
        }
        // 計算從當前時間到下次晚上10點的時間差
        private TimeSpan GetTimeUntilNext10PM()
        {
            DateTime now = DateTime.Now;
            DateTime next10PM = new DateTime(now.Year, now.Month, now.Day, 22, 0, 0); // 今天的22:00

            if (now > next10PM)
            {
                // 如果已經過了今天的10點，設置到明天的10點
                next10PM = next10PM.AddDays(1);
            }

            return next10PM - now;
        }
        private void OnDeleteDocFiles(object sender, ElapsedEventArgs e)
        {

            listBoxErr.Items.Clear();

            ClearWordTempFiles();//先清除暫存檔

            string folderPath = System.Windows.Forms.Application.StartupPath;

            try
            {
                //// 找到所有的.doc文件
                //var docFiles = Directory.EnumerateFiles(folderPath, "*.doc");
                // 找到所有 .doc 檔案（包含子資料夾）
                var docFiles = Directory.EnumerateFiles(folderPath, "*.doc", SearchOption.AllDirectories);

                foreach (string file in docFiles)
                {
                    // 刪除文件
                    File.Delete(file);
                }

                // 日志或顯示刪除結果
                int docNum = docFiles.Count();
                this.Invoke((MethodInvoker)delegate
                {
                    listBoxErr.Items.Add($"成功刪除了 {docNum} 個 .doc 文件 at {DateTime.Now}");
                    listBoxErr.TopIndex = listBoxErr.Items.Count - 1;
                });
            }
            catch (Exception ex)
            {
                // 處理異常
                this.Invoke((MethodInvoker)delegate
                {
                    listBoxErr.Items.Add("刪除文件時出錯: " + ex.Message);
                    listBoxErr.TopIndex = listBoxErr.Items.Count - 1;
                });
            }

            // 重設計時器到明天的10點
            SetDailyDeleteTimer();
        }
        // 清除所有 WINWORD.EXE 進程
        private void ClearWordTempFiles()
        {
            try
            {
                // 列舉所有名為 WINWORD 的進程
                var wordProcesses = Process.GetProcessesByName("WINWORD");

                // 如果有正在運行的 Word 進程，將其全部終止
                if (wordProcesses.Any())
                {
                    foreach (var process in wordProcesses)
                    {
                        process.Kill();
                        process.WaitForExit(); // 確保進程被完全終止
                    }
                    this.Invoke((MethodInvoker)delegate
                    {
                        listBoxErr.Items.Add($"{wordProcesses.Length} 個 Word 進程已經終止。清除成功");
                        listBoxErr.TopIndex = listBoxErr.Items.Count - 1;
                    });

                    //MessageBox.Show($"{wordProcesses.Length} 個 Word 進程已經終止。", "清除成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // MessageBox.Show("沒有運行中的 Word 進程。", "清除成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Invoke((MethodInvoker)delegate
                    {
                        listBoxErr.Items.Add("沒有運行中的 Word 進程。。清除成功");
                        listBoxErr.TopIndex = listBoxErr.Items.Count - 1;
                    });
                }
            }
            catch (Exception ex)
            {
                // 處理異常情況
                this.Invoke((MethodInvoker)delegate
                {
                    listBoxErr.Items.Add("清除 Word 進程時發生錯誤: " + ex.Message);
                    listBoxErr.TopIndex = listBoxErr.Items.Count - 1;
                });
                //MessageBox.Show("清除 Word 進程時發生錯誤: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private TimerParameters GetTimerParametersFromDatabase(int timerIndex)
        {
            TimerParameters parameters = null;
            SqlDataReader dr = DBA.executeParameterReader(timerIndex);
            while (dr.Read())
            {
                parameters = new TimerParameters
                {
                    C_System = dr.GetString(0),
                    Product = dr.GetString(1),
                    P_Type = dr.GetString(2),
                    PrintNum = dr.GetString(3)
                };
            }
            dr.Close();
            return parameters;
        }
        public class TimerParameters
        {
            public int TimerIndex { get; set; }
            public string C_System { get; set; }
            public string Product { get; set; }
            public string P_Type { get; set; }
            public string PrintNum { get; set; }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            backgroundWorker1.RunWorkerAsync();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            EmailSender.SendEmailAsync("", "", "HankChang@GGA.ASIA", "TEST" + "停止超過10秒", "中斷時間" + DateTime.Now.AddSeconds(-10).ToString());
        }
    }
}
