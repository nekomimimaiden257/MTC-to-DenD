using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LibUsbDotNet; // USBライブラリの名前空間を使用する
using LibUsbDotNet.Main; // USBライブラリの名前空間を使用する
using System.Threading;

//参考URL: https://www.ipentec.com/document/libusbdotnet-app-create 

// checkUSB 名前空間に、MTC to DenD の機能を押し込む
namespace checkUSB
{
    /// <summary>
    /// デザインフォームのコード。
    /// この1フォームで全部をやる。
    /// app.manifest で起動時に管理者権限を要求する。
    /// </summary>
    public partial class Form1 : Form
    {

        #region VariableSection
        public static UsbDevice MyUsbDevice; //USBデバイス自体を格納する
        // MTC P4-B7+非常 専用 、USB機器のベンダIDと機器IDがMTCごとに違うから、ほかのMTCに対応できない。
        public static UsbDeviceFinder MyUsbFinder = new UsbDeviceFinder(0x0AE4, 0x0101); // USB機器のベンダーIDとプロダクトIDを指定する。
        bool IsDisposing = false;//Disposeが実行中かどうかをチェック

        //なぜ配列で持たないのかは不明
        private string senddata = string.Empty;
        private string Pnotchsend = string.Empty;
        private string Bnotchsend = string.Empty;
        private string beforsend = string.Empty;
        private string senddata1 = string.Empty;
        private string senddata2 = string.Empty;
        private string senddata3 = string.Empty;
        private string senddata4 = string.Empty;
        private string senddata5 = string.Empty;
        private string senddata6 = string.Empty;
        private string senddata7 = string.Empty;
        private string senddata8 = string.Empty;
        private string senddata9 = string.Empty;
        private string senddata10 = string.Empty;
        private string senddata11 = string.Empty;
        private string sendbotton = string.Empty;
        int nownotch = 0;
        int nowbrake = 0;
        int bottonnum1befor = 0;
        int bottonnum10befor = 0;
        int bottonnum100befor = 0;
        int bottonnum1000befor = 0;
        int bottonnum10000befor = 0;
        int bottonnum100000befor = 0;
        int bottonnum1_befor = 0;
        int bottonnum10_befor = 0;
        int bottonnum100_befor = 0;
        int bottonnum1000_befor = 0;
        int bottonnum10000_befor = 0;
        int bottonnum100000_befor = 0;

        #endregion

        /// <summary>
        /// Form1のコンストラクタ
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            this.ControlBox = false;
        }

        /// <summary>
        /// フォーム読み込み時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            label1.Text = "切断中";
            //USBのベンダーIDとプロダクトIDをTextBoxコントロールに表示する。
            textBox1.Text = "0x" + Convert.ToString( MTCtoDenD.Properties.Settings.Default.UsbVenderID , 16).ToUpper();
            textBox2.Text = "0x" + Convert.ToString( MTCtoDenD.Properties.Settings.Default.UsbProductID , 16).ToUpper();

        }

        /// <summary>
        /// フォームを閉じるとき≒終了時
        /// USBデバイスと終了フラグを立てる
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Closing(object sender, EventArgs e)
        {
            MyUsbDevice = null;
            IsDisposing = true;
            button5.PerformClick();//ベンダーIDとプロダクトIDを保存する。
        }

        /// <summary>
        /// button1( 接続ボタン )をクリックしたら
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks>USBのベンダIDとプロダクトIDの変更を禁止する。</remarks>
        private void button1_Click(object sender, EventArgs e)
        {
            ErrorCode ec = ErrorCode.None;

            button5.Enabled = false;

            try
            {
                // MyUsbFinderで見つけた、USBデバイスを開いて、MyUsbDeviceに格納する。
                MyUsbDevice = UsbDevice.OpenUsbDevice(MyUsbFinder);

                // MyUsbDevice が見つからんのであれば、デバイスが見つからないとException吐いて抜ける。
                if (MyUsbDevice == null) throw new Exception("Device Not Found.");
                IUsbDevice wholeUsbDevice = MyUsbDevice as IUsbDevice; // MyUsbDeviceがIUsbDeviceとして扱えるなら、続行。
                if (!ReferenceEquals(wholeUsbDevice, null))
                {
                    IsDisposing = false;
                    wholeUsbDevice.SetConfiguration(1);
                    wholeUsbDevice.ClaimInterface(0);
                    label1.Text = "接続完了";
                }

            }
            // USBデバイス識別段階でトラブったら、Label1にエラーメッセージを残す。
            catch (Exception ex)
            {
                label1.Text = ec != ErrorCode.None ? ec + ":" : String.Empty + ex.Message;
            }

            //データ読み取り用のスレッドを立てる。
            Thread thread = new Thread(new ThreadStart(() => {
                while (!IsDisposing)
                {//Disposeが呼ばれるまで無限ループ
                    DataRead();
                }
            }));
            thread.Start();

            // データ送信用のすれどを立てる。
            Thread thread2 = new Thread(new ThreadStart(() => {
                while (!IsDisposing)
                {//Disposeが呼ばれるまで無限ループ
                    Senddata();
                }
            }));
            thread2.Start();

        }

        /// <summary>
        /// button3（ 切断ボタン ）デバイスだけを閉じて、アプリケーションは終了しない。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks>USBのベンダIDとプロダクトIDの変更を許可する</remarks>
        private void button3_Click(object sender, EventArgs e)
        {
            button5.Enabled = true;
            MyUsbDevice = null;
            IsDisposing = true;
            label1.Text = "切断中";
        }

        /// <summary>
        /// Button2（ 終了ボタン ）クリック時
        /// USBデバイスを解放して、終了する。
        /// ラベルに切断中を表示するが、多分表示完了前にアプリケーションが終了する。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks>USBのベンダIDとプロダクトIDの変更を許可する</remarks>
        private void button2_Click_1(object sender, EventArgs e)
        {
            button5.Enabled = true;
            MyUsbDevice = null;
            IsDisposing = true;
            label1.Text = "切断中";
            Application.Exit(); //アプリケーション終了

        }

        /// <summary>
        /// 受信 （  ）メソッド
        /// </summary>
        private void DataRead()
        {
            ErrorCode ec = ErrorCode.None;
            //リーダーを立てる。
            UsbEndpointReader reader = MyUsbDevice.OpenEndpointReader(ReadEndpointID.Ep01);
            byte[] readBuffer = new byte[8]; //データはバイト単位で来る。
            int bytesRead;
            ec = reader.Read(readBuffer, 100, out bytesRead);//データ受信
            String str = BitConverter.ToString(readBuffer);//読み取ったデータを文字列に変換する。
            //label2.Text = str;

            //マスコンのノッチ？
            switch (readBuffer[1])
            {
                // 逆転機=ニュートラル
                case 0x01:
                    //label3.Text = "EB";
                    nownotch = 0;
                    nowbrake = 0;
                    break;
                case 0x02:
                    //label3.Text = "B7";
                    nownotch = 0;
                    nowbrake = 0;
                    break;
                case 0x03:
                    //label3.Text = "B6";
                    nownotch = 0;
                    nowbrake = 0;
                    break;
                case 0x04:
                    //label3.Text = "B5";
                    nownotch = 0;
                    nowbrake = 0;
                    break;
                case 0x05:
                    //label3.Text = "B4";
                    nownotch = 0;
                    nowbrake = 0;
                    break;
                case 0x06:
                    //label3.Text = "B3";
                    nownotch = 0;
                    nowbrake = 0;
                    break;
                case 0x07:
                    //label3.Text = "B2";
                    nownotch = 0;
                    nowbrake = 0;
                    break;
                case 0x08:
                    //label3.Text = "B1";
                    nownotch = 0;
                    nowbrake = 0;
                    break;
                case 0x09:
                    //label3.Text = "N";
                    nownotch = 0;
                    nowbrake = 0;
                    break;
                case 0x0A:
                    //label3.Text = "p1";
                    nownotch = 0;
                    nowbrake = 0;
                    break;
                case 0x0B:
                    //label3.Text = "p2";
                    nownotch = 0;
                    nowbrake = 0;
                    break;
                case 0x0C:
                    //label3.Text = "p3";
                    nownotch = 0;
                    nowbrake = 0;
                    break;
                case 0x0D:
                    //label3.Text = "p4";
                    nownotch = 0;
                    nowbrake = 0;
                    break;
                //逆転機=前
                case 0x41:
                    //label3.Text = "EB";
                    nownotch = 0;
                    nowbrake = 9;
                    break;
                case 0x42:
                    //label3.Text = "B7";
                    nownotch = 0;
                    nowbrake = 7;
                    break;
                case 0x43:
                    //label3.Text = "B6";
                    nownotch = 0;
                    nowbrake = 6;
                    break;
                case 0x44:
                    //label3.Text = "B5";
                    nownotch = 0;
                    nowbrake = 5;
                    break;
                case 0x45:
                    //label3.Text = "B4";
                    nownotch = 0;
                    nowbrake = 4;
                    break;
                case 0x46:
                    //label3.Text = "B3";
                    nownotch = 0;
                    nowbrake = 3;
                    break;
                case 0x47:
                    //label3.Text = "B2";
                    nownotch = 0;
                    nowbrake = 2;
                    break;
                case 0x48:
                    //label3.Text = "B1";
                    nownotch = 0;
                    nowbrake = 1;
                    break;
                case 0x49:
                    //label3.Text = "N";
                    nownotch = 0;
                    nowbrake = 0;
                    break;
                case 0x4A:
                    //label3.Text = "p1";
                    nownotch = 1;
                    nowbrake = 0;
                    break;
                case 0x4B:
                    //label3.Text = "p2";
                    nownotch = 2;
                    nowbrake = 0;
                    break;
                case 0x4C:
                    //label3.Text = "p3";
                    nownotch = 3;
                    nowbrake = 0;
                    break;
                case 0x4D:
                    //label3.Text = "p4";
                    nownotch = 4;
                    nowbrake = 0;
                    break;
                case 0x81:
                    //label3.Text = "EB";
                    nownotch = 0;
                    nowbrake = 9;
                    break;
                case 0x82:
                    //label3.Text = "B7";
                    nownotch = 0;
                    nowbrake = 7;
                    break;
                case 0x83:
                    //label3.Text = "B6";
                    nownotch = 0;
                    nowbrake = 6;
                    break;
                case 0x84:
                    //label3.Text = "B5";
                    nownotch = 0;
                    nowbrake = 5;
                    break;
                case 0x85:
                    //label3.Text = "B4";
                    nownotch = 0;
                    nowbrake = 4;
                    break;
                case 0x86:
                    //label3.Text = "B3";
                    nownotch = 0;
                    nowbrake = 3;
                    break;
                case 0x87:
                    //label3.Text = "B2";
                    nownotch = 0;
                    nowbrake = 2;
                    break;
                case 0x88:
                    //label3.Text = "B1";
                    nownotch = 0;
                    nowbrake = 1;
                    break;
                case 0x89:
                    //label3.Text = "N";
                    nownotch = 0;
                    nowbrake = 0;
                    break;
                case 0x8A:
                    //label3.Text = "p1";
                    nownotch = 1;
                    nowbrake = 0;
                    break;
                case 0x8B:
                    //label3.Text = "p2";
                    nownotch = 2;
                    nowbrake = 0;
                    break;
                case 0x8C:
                    //label3.Text = "p3";
                    nownotch = 3;
                    nowbrake = 0;
                    break;
                case 0x8D:
                    //label3.Text = "p4";
                    nownotch = 4;
                    nowbrake = 0;
                    break;
            }

            

            
            //スロットルニュートラル
            if (nownotch == 0)
            {
                Pnotchsend = null;
            }
            else if (nownotch > 0) //力行＜りっこう＞ノッチ
            {
                Pnotchsend = "A";
            }
            if (nowbrake == 0)//ブレーキ緩解
            {
                Bnotchsend = null;
            }
            else if (nowbrake > 0 && nowbrake != 9) //常用ブレーキ
            {
                Bnotchsend = "S";
            }
            else if (nowbrake == 9) //非常ブレーキ
            {
                Bnotchsend = "D";
            }
            

            //ボタンの処理
            string bottonbin_ = Convert.ToString(readBuffer[3], 2);
            int bottonnum_ = Convert.ToInt32(bottonbin_);

            //start・"home"
            int bottonnum1_ = bottonnum_ - (bottonnum_ / 10) * 10;
            if (bottonnum1_ != bottonnum1_befor)
            {
                if (bottonnum1_ % 2 == 1)
                {
                    senddata1 = "{home}";
                }
                else
                {
                    senddata1 = null;
                }
            }
            else
            {
                senddata1 = null;
            }
            bottonnum1_befor = bottonnum1_;

            //select・"esc"
            int bottonnum10_ = (bottonnum_ - (bottonnum_ / 100) * 100) / 10;
            if (bottonnum10_ != bottonnum10_befor)
            {
                if (bottonnum10_ % 2 == 1)
                {
                    senddata2 = "{esc}";
                }
                else
                {
                    senddata2 = null;
                }
            }
            else
            {
                senddata2 = null;
            }
            bottonnum10_befor = bottonnum10_;

            //↑
            int bottonnum100_ = (bottonnum_ - (bottonnum_ / 1000) * 1000) / 100;
            if (bottonnum100_ != bottonnum100_befor)
            {
                if (bottonnum100_ % 2 == 1)
                {
                    senddata3 = "{UP}";
                }
                else
                {
                    senddata3 = null;
                }
            }
            bottonnum100_befor = bottonnum100_;

            //↓
            int bottonnum1000_ = (bottonnum_ - (bottonnum_ / 10000) * 10000) / 1000;
            if (bottonnum1000_ != bottonnum1000_befor)
            {
                if (bottonnum1000_ % 2 == 1)
                {
                    senddata4 = "{DOWN}";
                }
                else
                {
                    senddata4 = null;
                }
            }
            bottonnum1000_befor = bottonnum1000_;


            //←
            int bottonnum10000_ = (bottonnum_ - (bottonnum_ / 100000) * 100000) / 10000;
            if (bottonnum10000_ != bottonnum10000_befor)
            {
                if (bottonnum10000_ % 2 == 1)
                {
                    senddata5 = "{LEFT}";
                }
                else
                {
                    senddata5 = null;
                }
            }
            bottonnum10000_befor = bottonnum10000_;

            //→
            int bottonnum100000_ = (bottonnum_ - (bottonnum_ / 1000000) * 1000000) / 100000;
            if (bottonnum100000_ != bottonnum100000_befor)
            {
                if (bottonnum100000_ % 2 == 1)
                {
                    senddata6 = "{RIGHT}";
                }
                else
                {
                    senddata6 = null;
                }
            }
            bottonnum100000_befor = bottonnum100000_;


            //ボタンの処理
            string bottonbin = Convert.ToString(readBuffer[2], 2);
                int bottonnum = Convert.ToInt32(bottonbin);

            //S・未実装
            int bottonnum1 = bottonnum - (bottonnum / 10) * 10;
            if (bottonnum1 != bottonnum1befor)
            {
                if (bottonnum1 % 2 == 1)
                {
                    senddata7 = null;
                }
                else
                {
                    senddata7 = null;
                }
            }
            else
            {
                senddata7 = null;
            }
            bottonnum1befor = bottonnum1;

            //D・"W"
            int bottonnum10 = (bottonnum - (bottonnum / 100) * 100) / 10;
            if (bottonnum10 != bottonnum10befor)
            {
                if (bottonnum10 % 2 == 1)
                {
                    senddata8 = "W";
                }
                else
                {
                    senddata8 = null;
                }
            }
            else
            {
                senddata8 = null;
            }
            bottonnum10befor = bottonnum10;

            //A・"Q"
            int bottonnum100 = (bottonnum - (bottonnum / 1000) * 1000) / 100;
            if (bottonnum100 != bottonnum100befor)
            {
                if (bottonnum100 % 2 == 1)
                {
                    senddata9 = "Q";
                }
                else
                {
                    senddata9 = null;
                }
            }
            else
            {
                senddata9 = null;
            }
            bottonnum100befor = bottonnum100;

            //A深押し
            int bottonnum1000 = (bottonnum - (bottonnum / 10000) * 10000) / 1000;
            if (bottonnum1000 != bottonnum1000befor)
            {
            }
            bottonnum1000befor = bottonnum1000;


            //B・"エンター"
            int bottonnum10000 = (bottonnum - (bottonnum / 100000) * 100000) / 10000;
            if (bottonnum10000 != bottonnum10000befor)
            {
                if (bottonnum10000 % 2 == 1)
                {
                    senddata10 = "{ENTER}";
                }
                else
                {
                    senddata10 = null;
                }
            }
            else
            {
                senddata10 = null;
            }
            bottonnum10000befor = bottonnum10000;

            //C・未実装
            int bottonnum100000 = (bottonnum - (bottonnum / 1000000) * 1000000) / 100000;
            if (bottonnum100000 != bottonnum100000befor)
            {
                if (bottonnum100000 % 2 == 1)
                {
                    senddata11 = null;
                }
                else
                {
                    senddata11 = null;
                }
            }
            else
            {
                senddata11 = null;
            }
            bottonnum100000befor = bottonnum100000;





            sendbotton = senddata1 + senddata2 + senddata3 + senddata4 + senddata5 + senddata6 + senddata7 + senddata8 + senddata9 + senddata10 + senddata11;
            senddata = Pnotchsend + Bnotchsend;


            // 軽量化有効なら、Wait入れてCPU休ませつつ、メッセージ消化させる
            // コストとして50ms 消費する( 60fps（16.67ms）基準なら3フレーム遅延する)
            // Windowsのタイマーが50msでスレッド切り替えなので
            if (checkBox1.Checked)
            {
                SendKeys.SendWait(sendbotton);
                SendKeys.Flush();//DoEvents()と等価？
                Application.DoEvents();
                Thread.Sleep(1); //1msだけCPUを開放するが、コンテキストスイッチさせない。しかし、実際には、16msは消費する。
            }
            else
            {
                // 軽量化無効であれば従来通り
                SendKeys.SendWait(sendbotton);
            }

        }


        /// <summary>
        /// 送信（  →  ）メソッド
        /// </summary>
        /// <remarks>
        /// もしかして、ノッチの文字情報を電車でDへ送ってるだけ？
        /// </remarks>
        private void Senddata()
        {
            //SendKeys.SendWait();

            // 軽量化有効なら、Wait入れてCPU休ませつつ、メッセージ消化させる
            // コストとして50ms 消費する( 60fps（16.67ms）基準なら3フレーム遅延する)
            // Windowsのデフォルトタイマーが50msでスレッド切り替えの分解能なので
            if (checkBox1.Checked)
            {
                SendKeys.SendWait(senddata);
                SendKeys.Flush();
                Application.DoEvents();//DoEvents()と等価？
                Thread.Sleep(1); //1msだけCPUを開放するが、コンテキストスイッチさせない。しかし、実際には、16msは消費する。
            }
            else 
            {
                // 軽量化無効であれば、従来通り
                SendKeys.SendWait(senddata);
            }
        }

        /// <summary>
        /// button4をクリックしたら、電車でD SSを起動する。同じフォルダーに電車でD SSの実行ファイルがある事！
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            const string progname = @"電車でＤ ShiningStage.exe"; //電車でD ShingStageの実行ファイル名を書き換え不能で宣言
            if ( System.IO.File.Exists(progname) ) {
                // 電車でDの実行ファイルがカレントディレクトリに存在する -> 起動
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo.FileName = progname;
                proc.StartInfo.UseShellExecute = true; // シェルモードON（UnityのChsarpだから、falseでもあがる？）
                // 管理者として実行する設定
                proc.StartInfo.Verb = "RunAs";
                try
                {
                    proc.Start();//電車でD ShingStageプロセスの起動
                }
                catch (Exception Ex) {
                    label1.Text = Ex.ToString();
                    MessageBox.Show("電車での実行ファイル起動失敗。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                //起動させたら野放しで、MTC to DenD では関知しない。
            }
            else {
                // 電車でDの実行ファイルがない -> 何もできない。
                MessageBox.Show("電車でD ShingStage が見つかりませんでした。", this.Text, MessageBoxButtons.OK , MessageBoxIcon.Error );
            
            }
        }

        /// <summary>
        /// button5をクリックしたら、入力されたUSBベンダーIDとプロダクトIDの値を設定に保存する。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                //ベンダーIDとプロダクトID（16進数）をIntに変換して、値をセットする。
                MTCtoDenD.Properties.Settings.Default.UsbVenderID = Convert.ToInt32(textBox1.Text, 16);
                MTCtoDenD.Properties.Settings.Default.UsbProductID = Convert.ToInt32(textBox2.Text, 16);

            }
            catch (FormatException Fex)
            {
                label1.Text = Fex.ToString();
                MessageBox.Show("ベンダーIDかプロダクトIDの数値形式がおかしい。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception Ex) {
                label1.Text = Ex.ToString();
                MessageBox.Show("ベンダーIDかプロダクトIDの変換失敗。" , this.Text , MessageBoxButtons.OK , MessageBoxIcon.Error);
            }
            // この時点でセーブする。
            MTCtoDenD.Properties.Settings.Default.Save();
            //USBのベンダベンダIDとプロダクトIDをUSBファインダーに登録する
            MyUsbFinder = new UsbDeviceFinder( Convert.ToInt32( textBox1.Text ) ,Convert.ToInt32( textBox2.Text ) );
        }
    }

}   