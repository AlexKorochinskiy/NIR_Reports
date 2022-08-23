using ClosedXML.Excel;
using DataModels;
using LinqToDB;
using LinqToDB.Data;
using mcOMRON;
using MomentSharp;
using OpcLabs.EasyOpc.DataAccess;
using OpcLabs.EasyOpc.OperationModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OPCtoOmron
{
    public partial class MainForm : MaterialSkin.Controls.MaterialForm
    {
        private NLog.Logger logger;
        private volatile OmronPLC plc;
        private int ErrorCounter; //Счетчик ошибок связи с PLC
        private bool IsRestarting = false;
        private Cache memoryCache;
        private DataRetriever retriever;
        private string connString;
        private Cache memoryCacheJournal;
        private DataRetriever retrieverJournal;
        private List<Operator> operators;
        private List<Calibr> calibrs;

        public MainForm()
        {
            InitializeComponent();
            StackedHeaderDecorator objREnderer = new StackedHeaderDecorator(dataGridView1);
            StackedHeaderDecorator objREnderer1 = new StackedHeaderDecorator(dataGridView3);
            Program.materialSkinManager.AddFormToManage(this);
            typeof(DataGridView).InvokeMember("DoubleBuffered", BindingFlags.NonPublic |
                BindingFlags.Instance | BindingFlags.SetProperty, null,
                dataGridView1, new object[] { true });
            LinqToDB.Common.Configuration.Linq.AllowMultipleQuery = true;
            dataGridView1.DefaultCellStyle.ForeColor = Color.Black;
            dataGridView2.DefaultCellStyle.ForeColor = Color.Black;
            dataGridView3.DefaultCellStyle.ForeColor = Color.Black;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason != CloseReason.WindowsShutDown) //Если это не перезагрузка или выключение Windows
            {
                e.Cancel = true;
                if (IsRestarting || MessageBox.Show("Вы действительно хотите закрыть приложение?",
                        this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    if (backgroundWorker1.IsBusy)
                        backgroundWorker1.CancelAsync();
                    else
                    {
                        if (plc != null && plc.Connected)
                            plc.Close();
                        logger.Info("Приложение закрыто пользователем.");
                        e.Cancel = false;
                    }
                }
            }
        }

        private async void MainForm_Shown(object sender, EventArgs e)
        {
            await Task.Factory.StartNew(() =>
            {
                try
                {
                    logger = NLog.LogManager.GetCurrentClassLogger();
                    logger.Info("Приложение запущено.");

                    //materialLabel3.SafeInvoke(p => p.Font = new Font(materialLabel3.Font.FontFamily, 20));
                    //materialLabel4.SafeInvoke(p => p.Font = new Font(materialLabel4.Font.FontFamily, 20));
                    //materialLabel5.SafeInvoke(p => p.Font = new Font(materialLabel5.Font.FontFamily, 20));

                    using (NirDBDB nir = new NirDBDB())
                    {
                        operators = nir.Operators.ToList();
                        calibrs = nir.Calibrs.ToList();
                    }

                    dataGridView1.SafeInvoke(p => p.RowTemplate.Height = 30);
                    dataGridView2.SafeInvoke(p => p.RowTemplate.Height = 30);

                    dataGridView1.SafeInvoke(p => (p.Columns[3] as DataGridViewComboBoxColumn).DataSource = calibrs);
                    dataGridView1.SafeInvoke(p => (p.Columns[3] as DataGridViewComboBoxColumn).ValueMember = "ID");
                    dataGridView1.SafeInvoke(p => (p.Columns[3] as DataGridViewComboBoxColumn).DisplayMember = "Name");

                    dataGridView1.SafeInvoke(p => (p.Columns[8] as DataGridViewComboBoxColumn).DataSource = operators);
                    dataGridView1.SafeInvoke(p => (p.Columns[8] as DataGridViewComboBoxColumn).ValueMember = "ID");
                    dataGridView1.SafeInvoke(p => (p.Columns[8] as DataGridViewComboBoxColumn).DisplayMember = "OperatorColumn");

                    connString = NativeMetods.ReadINIString("SQL", "SQLConnectionString");

                    //для подключения к PLC
                    plc = new OmronPLC(mcOMRON.TransportType.Tcp);
                    tcpFINSCommand tcpCommand = (tcpFINSCommand)plc.FinsCommand;
                    tcpCommand.SetTCPParams(IPAddress.Parse(NativeMetods.ReadINIString("PLC", "IP")),
                        NativeMetods.ReadINIInt("PLC", "Port"), NativeMetods.ReadINIInt("PLC", "Timeout"));

                    if (plc.Connect())
                    {
                        logger.Info("Подключение к Omron PLC успешно произведено.");
                        statusStrip1.SafeInvoke(p => toolStripStatusLabel1.BackColor = Color.Green);
                    }
                    else
                    {
                        logger.Error("Ошибка подключения к Omron PLC {0}", plc.LastError);
                        ErrorCounter = 4;
                    }

                    retriever = new DataRetriever(connString, "NIR_Params")
                    {
                        ColumnsToAddList = "*",
                        ColumnToSortBy = "ID",
                        Filter = string.Empty
                    };

                    retrieverJournal = new DataRetriever(connString, "Journal")
                    {
                        ColumnsToAddList = "*",
                        ColumnToSortBy = "ID",
                        Filter = string.Empty
                    };

                    backgroundWorker1.RunWorkerAsync();

                    dataGridView1.Rows.Clear();
                    retriever.Filter = string.Empty;
                    int rowCount = retriever.RowCount;
                    memoryCache = new Cache(retriever, 30);
                    dataGridView1.SafeInvoke(p => { p.RowCount = rowCount; });
                }
                catch (OpcException ex)
                {
                    logger.Error("[{0}] {1}", ex.LineNumber(), ex.InnerException.Message);
                }
                catch (Exception ex)
                {
                    logger.Error("[{0}] {1}", ex.LineNumber(), ex.Message);
                }
            });
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            foreach (var item in statusStrip1.Items)
            {
                if (item is ToolStripStatusLabel)
                {
                    ToolStripStatusLabel toolStripStatusLabel = item as ToolStripStatusLabel;
                    if (toolStripStatusLabel.BackColor == Color.Red)
                        toolStripStatusLabel.BackColor = Color.DarkRed;
                    else if (toolStripStatusLabel.BackColor == Color.DarkRed)
                        toolStripStatusLabel.BackColor = Color.Red;

                    if (toolStripStatusLabel.BackColor == Color.Green)
                        toolStripStatusLabel.BackColor = Color.DarkGreen;
                    else if (toolStripStatusLabel.BackColor == Color.DarkGreen)
                        toolStripStatusLabel.BackColor = Color.Green;
                }
            }
        }

        private bool StopedWithError = false;
        private void BackgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            Stopwatch stopwatch = null;
            FileStream fs = null;
            StreamReader f = null;
            NirDBDB nirDB = null;
            NirParam nirParam = null;
            Journal journal = null;

            string SqlDbSize = "SELECT LTRIM(STR((CONVERT(dec(15, 2), dbsize) + CONVERT(dec(15, 2), logsize)) * 8192 / 1048576, 15, 2)) DB_Size " +
                "FROM " +
                "( " +
                " SELECT SUM(CONVERT(BIGINT, CASE WHEN status & 64 = 0 THEN size ELSE 0 END)) dbsize " +
                " , SUM(CONVERT(BIGINT, CASE WHEN status & 64 <> 0 THEN size ELSE 0 END)) logsize " +
                " FROM dbo.sysfiles " +
                ")big";

            string SqlVersion = "SELECT @@VERSION ";

            try
            {
                int timeout = NativeMetods.ReadINIInt("General", "Timeout");
                string machineName = NativeMetods.ReadINIString("OPC", "MachineName");
                string serverName = NativeMetods.ReadINIString("OPC", "Server");
                string lastTask = NativeMetods.ReadINIString("General", "LastTask");

                nirDB = new NirDBDB();
                string path = NativeMetods.ReadINIString("ArchiveVspFile", "Path");

                string line = null;

                fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                f = new StreamReader(fs, Encoding.Default);

                stopwatch = new Stopwatch();

                int FossWatchdog = 0;
                int firstCalibrationId = nirDB.Calibrs.FirstOrDefault().ID;
                int recipeType = 0;
                int recipeSubType = 0;
                float moi = 0;
                Mixing mixing = Mixing.Stoped;
                int Otves = 0;
                ushort WaterPercent = 0;
                bool old_unload = false;
                string[] sds = NativeMetods.ReadINIString("SDS", "Availble").Split(';');
                int productCodeNum = 0;

                journal = nirDB.Journals.OrderByDescending(p => p.DateTime).FirstOrDefault();
                if (journal == null)
                {
                    //Если база данных пустая
                    journal = new Journal();
                }
                else
                {
                    recipeType = journal.RecipeType;
                    recipeSubType = journal.RecipeSubType;
                    Otves = journal.NumberOtves;
                }

                long pos = NativeMetods.ReadINIlong("ArchiveVspFile", "LastPosition");
                if (f.Length() < pos)
                {
                    logger.Warn("Файл архива VSP был изменен.");
                    if (MessageBox.Show("Файл архива VSP5000 был изменен. Продолжить чтение из этого файла?",
                        this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    {
                        StopedWithError = true;
                        return;
                    }
                    else
                    {
                        f.SetPosition(0);
                        logger.Info("Пользователь продолжил чтение файла архива сначала при несовпадении.");
                    }
                }
                else
                {
                    f.SetPosition(pos);
                }

                while (!worker.CancellationPending)
                {
                    //Проверяю размер базы 
                    bool Express = nirDB.Execute<string>(SqlVersion).Contains("Express");
                    double size = nirDB.Execute<double>(SqlDbSize);

                    if (Express && size > 9900 && (stopwatch.ElapsedMilliseconds >= 600000 || !stopwatch.IsRunning))
                    {
                        MessageBox.Show("ВНИМАНИЕ!!! Размер базы данных подходит к максимальному значению." + Environment.NewLine
                                + "После достижения 10ГБ дальнейшие записи производиться не будут!!! Необходимо очистить БД!!!");
                        stopwatch.Restart();
                    }

                    if (Express && size > 9600)
                    {
                        logger.Warn("ВНИМАНИЕ!!! Размер базы данных подходит к максимальному значению ({0}). "
                                + "После достижения 10ГБ дальнейшие записи производиться не будут!!! Необходимо очистить БД!!!", size);
                    }

                    //Переподключение к PLC при ошибках связи
                    if (ErrorCounter > 3)
                    {
                        if (plc.Connected)
                            plc.Close();
                        if (plc.Connect())
                        {
                            ErrorCounter = 0;
                            logger.Info("Переподключение к Omron PLC успешно произведено.");
                            statusStrip1.SafeInvoke(p => toolStripStatusLabel1.BackColor = Color.Green);
                        }
                        else
                        {
                            statusStrip1.SafeInvoke(p => toolStripStatusLabel1.BackColor = Color.Red);
                            logger.Error("Ошибка подключения к Omron PLC {0}", plc.LastError);
                            int tmp_timeout1 = timeout;
                            while (tmp_timeout1 >= 0 && !worker.CancellationPending)
                            {
                                tmp_timeout1--;
                                Thread.Sleep(1);
                            }
                            continue;
                        }
                    }

                    //Проверяю подключение к OPC и пишу Watchdog
                    try
                    {
                        easyDAClient1.WriteItemValue(machineName, serverName, "FOSS.ProFoss.Controller.WatchdogCounter", FossWatchdog);
                        FossWatchdog++;
                    }
                    catch (OpcException ex)
                    {
                        logger.Error("[{0}] Ошибка в потоке обмена - {1}", ex.LineNumber(), ex.InnerException.Message);
                        int tmp_timeout1 = timeout;
                        while (tmp_timeout1 >= 0 && !worker.CancellationPending)
                        {
                            tmp_timeout1--;
                            Thread.Sleep(1);
                        }
                        continue;
                    }

                    line = f.ReadLine();

                    if (line != null)
                    {
                        logger.Debug(line);
                        TotalPartOfString mes = GetTotalPartOfString(line);
                        if (mes.RecipeNum.Length > 5 && mes.RecipeNum[4] == '-' &&
                            int.TryParse(mes.RecipeNum.Substring(0, 2), out recipeType)
                            && int.TryParse(mes.RecipeNum.Substring(2, 2), out recipeSubType))
                        {
                            if (sds.Contains(mes.DeviceName) && !mes.Start)
                            {
                                if (!nirDB.NirParams.Any(p => p.RecipeType == recipeType && p.RecipeSubType == recipeSubType))
                                {
                                    //Если появилась новая строка вставляю строку в таблицу параметров
                                    int operators = nirDB.Operators.OrderBy(p => p.ID).Select(p=> p.ID).FirstOrDefault();
                                    if(operators <= 0)
                                    {
                                        MessageBox.Show("В базу данных не добавлен ни одна фамилия оператора. "
                                            + "Выберите настройки и внесите данные. После этого перезапустите программу.", 
                                            Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        return;
                                    }

                                    nirDB.NirParams
                                        .Value(p => p.RecipeType, recipeType)
                                        .Value(p => p.RecipeSubType, recipeSubType)
                                        .Value(p => p.Calibration, firstCalibrationId)
                                        .Value(p => p.MaxPercent, 0)
                                        .Value(p => p.MinPercent, 0)
                                        .Value(p => p.Operator, operators)
                                        .Value(p => p.StartPercent, 0)
                                        .Value(p => p.TargetPercent, 0)
                                        .Value(p => p.EnableCorrection, false)
                                        .Insert();

                                    RefreshParams();
                                }
                                
                                DataGridViewRow row = dataGridView1.Rows.Cast<DataGridViewRow>()
                                    .Where(p=> Convert.ToInt32(p.Cells[1].Value) == recipeType && Convert.ToInt32(p.Cells[2].Value) == recipeSubType)
                                    .FirstOrDefault();
                                if(row != null)
                                {
                                    dataGridView1.SafeInvoke(p=> p.ClearSelection());
                                    dataGridView1.SafeInvoke(p=> p.FirstDisplayedScrollingRowIndex = row.Index);
                                    dataGridView1.SafeInvoke(p=> p.Rows[row.Index].Selected = true);
                                }
                            }
                            else if (sds.Contains(mes.DeviceName) && mes.Start)
                            {
                                //Если был старт рецепта, то обновляю старт
                                nirParam = nirDB.NirParams.Where(p => p.RecipeType == recipeType && p.RecipeSubType == recipeSubType).FirstOrDefault();
                                productCodeNum = nirDB.Calibrs.Where(p => p.ID == nirParam.Calibration).Select(p => p.Code).FirstOrDefault();
                                Otves = 0;
                                //NativeMetods.WriteINI("General", "Otves", Otves);

                                //Записываю стартовый процент ввода воды
                                //if (plc.WriteDM(1234, (ushort)nirParam.StartPercent))
                                //    ErrorCounter = 0;
                                //else
                                //{
                                //    ErrorCounter++;
                                //    logger.Warn("Ошибка чтения из Omron PLC - {0}", plc.LastError);
                                //}
                            }
                        }

                        pos = f.GetPosition();
                        NativeMetods.WriteINI("ArchiveVspFile", "LastPosition", pos);
                    }

                    //сухое смешивание
                    bool dry_mixing = false;
                    if (plc.ReadDM_Bit(18485, 0, ref dry_mixing))
                        ErrorCounter = 0;
                    else
                    {
                        ErrorCounter++;
                        logger.Warn("Ошибка чтения из Omron PLC - {0}", plc.LastError);
                    }

                    //влажное смешивание
                    bool wet_mixing = false;
                    if (plc.ReadDM_Bit(18485, 1, ref wet_mixing))
                        ErrorCounter = 0;
                    else
                    {
                        ErrorCounter++;
                        logger.Warn("Ошибка чтения из Omron PLC - {0}", plc.LastError);
                    }

                    //Выгрузка
                    bool unload = false;
                    if (plc.ReadDM_Bit(18485, 6, ref unload))
                        ErrorCounter = 0;
                    else
                    {
                        ErrorCounter++;
                        logger.Warn("Ошибка чтения из Omron PLC - {0}", plc.LastError);
                    }

                    //Пересчет влажности и внесение корректировки
                    if (unload && unload != old_unload)
                    {
                        Otves++;

                        if (plc.ReadDM(18490, ref WaterPercent))
                            ErrorCounter = 0;
                        else
                        {
                            ErrorCounter++;
                            logger.Warn("Ошибка чтения из Omron PLC - {0}", plc.LastError);
                        }

                        //nirDB.Journals
                        //    .Value(p => p.DateTime, DateTime.Now)
                        //    .Value(p => p.Calibration, nirParam.Calibration)
                        //    .Value(p => p.MoiPercent, moi)
                        //    .Value(p => p.NumberOtves, Otves)
                        //    .Value(p => p.Operator, nirParam.Operator)
                        //    .Value(p => p.RecipeType, nirParam.RecipeType)
                        //    .Value(p => p.RecipeSubType, nirParam.RecipeSubType)
                        //    .Value(p => p.WaterPercent, WaterPercent)
                        //    .Value(p=> p.EnableCorrection, nirParam.EnableCorrection)
                        //    .Insert();

                        //RefreshJournal();

                        switch (mixing)
                        {
                            case Mixing.Dry: //Сухое смешивание





                                break;
                            case Mixing.Wet: //Влажное смешивание
                                WaterPercent = (ushort)(WaterPercent * moi / nirParam.TargetPercent);
                                logger.Debug("Расчетный процент = {0}", WaterPercent);
                                if (nirParam.EnableCorrection)
                                {
                                    //Если коррекция разрешена
                                    //if (plc.WriteDM(1234, WaterPercent))
                                    //    ErrorCounter = 0;
                                    //else
                                    //{
                                    //    ErrorCounter++;
                                    //    logger.Warn("Ошибка чтения из Omron PLC - {0}", plc.LastError);
                                    //}
                                }
                                break;
                        }
                    }

                    old_unload = unload;

                    float val0 = 0f;
                    byte[] arr0 = BitConverter.GetBytes(val0).SwapBytes();

                    if (dry_mixing) //Сухое смешивание
                    {
                        string prodCodeName = nirDB.Calibrs.Where(p => p.Code == productCodeNum).Select(p => p.Name).FirstOrDefault();
                        materialLabel3.SafeInvoke(p => p.Text = prodCodeName);
                        materialLabel4.SafeInvoke(p => p.Text = "Сухое смешивание");
                        mixing = Mixing.Dry;

                        easyDAClient1.WriteItemValue(machineName, serverName, "FOSS.ProFoss.Controller.ProductCodeN", productCodeNum);
                        Thread.Sleep(100);
                        easyDAClient1.WriteItemValue(machineName, serverName, "FOSS.ProFoss.Controller.StartMeasuring", 1);
                        Thread.Sleep(100);

                        object Value = easyDAClient1.ReadItemValue(machineName, serverName, "FOSS.ProFoss.Sample.Parameters.Moi.Result");
                        moi = Convert.ToSingle(Value);
                        materialLabel5.SafeInvoke(p => p.Text = moi.ToString("000.00"));

                        byte[] arr = BitConverter.GetBytes(moi).SwapBytes();

                        if (plc.finsMemoryAreadWrite(MemoryArea.DM, 18486, 0, 2, arr))
                            ErrorCounter = 0;
                        else
                        {
                            ErrorCounter++;
                            logger.Warn("Ошибка записи в Omron PLC - {0}", plc.LastError);
                        }

                        if (plc.finsMemoryAreadWrite(MemoryArea.DM, 18488, 0, 2, arr0))
                            ErrorCounter = 0;
                        else
                        {
                            ErrorCounter++;
                            logger.Warn("Ошибка записи в Omron PLC - {0}", plc.LastError);
                        }
                    }
                    else if (wet_mixing) //Влажное смешивание
                    {
                        string prodCodeName = nirDB.Calibrs.Where(p => p.Code == productCodeNum).Select(p => p.Name).FirstOrDefault();
                        materialLabel3.SafeInvoke(p => p.Text = prodCodeName);

                        materialLabel4.SafeInvoke(p => p.Text = "Влажное смешивание");
                        mixing = Mixing.Wet;

                        easyDAClient1.WriteItemValue(machineName, serverName, "FOSS.ProFoss.Controller.ProductCodeN", productCodeNum);
                        Thread.Sleep(100);
                        easyDAClient1.WriteItemValue(machineName, serverName, "FOSS.ProFoss.Controller.StartMeasuring", 1);
                        Thread.Sleep(100);

                        object Value = easyDAClient1.ReadItemValue(machineName, serverName, "FOSS.ProFoss.Sample.Parameters.Moi.Result");
                        moi = Convert.ToSingle(Value);
                        materialLabel5.SafeInvoke(p => p.Text = moi.ToString("000.00"));

                        byte[] arr = BitConverter.GetBytes(moi).SwapBytes();

                        if (plc.finsMemoryAreadWrite(MemoryArea.DM, 18488, 0, 2, arr))
                            ErrorCounter = 0;
                        else
                        {
                            ErrorCounter++;
                            logger.Warn("Ошибка записи в Omron PLC - {0}", plc.LastError);
                        }

                        if (plc.finsMemoryAreadWrite(MemoryArea.DM, 18486, 0, 2, arr0))
                            ErrorCounter = 0;
                        else
                        {
                            ErrorCounter++;
                            logger.Warn("Ошибка записи в Omron PLC - {0}", plc.LastError);
                        }

                        //Выход значения влажности за пределы
                        if (nirParam != null && moi > nirParam.MaxPercent)
                        {
                            logger.Warn("Влажность больше MaxValue.");
                            if (nirParam.EnableCorrection)
                            {
                                if (plc.WriteDM_Bit(18485, 3, true))
                                    ErrorCounter = 0;
                                else
                                {
                                    ErrorCounter++;
                                    logger.Warn("Ошибка чтения из Omron PLC - {0}", plc.LastError);
                                }
                            }
                        }
                        else if (nirParam != null && moi < nirParam.MinPercent)
                        {
                            logger.Warn("Влажность меньше MinValue.");

                            if (nirParam.EnableCorrection)
                            {
                                if (plc.WriteDM_Bit(18485, 4, true))
                                    ErrorCounter = 0;
                                else
                                {
                                    ErrorCounter++;
                                    logger.Warn("Ошибка чтения из Omron PLC - {0}", plc.LastError);
                                }
                            }
                        }
                        else
                        {
                            if (nirParam != null && nirParam.EnableCorrection)
                            {
                                if (plc.WriteDM_Bit(18485, 3, false))
                                    ErrorCounter = 0;
                                else
                                {
                                    ErrorCounter++;
                                    logger.Warn("Ошибка чтения из Omron PLC - {0}", plc.LastError);
                                }

                                if (plc.WriteDM_Bit(18485, 4, false))
                                    ErrorCounter = 0;
                                else
                                {
                                    ErrorCounter++;
                                    logger.Warn("Ошибка чтения из Omron PLC - {0}", plc.LastError);
                                }
                            }
                        }
                    }
                    else
                    {
                        materialLabel3.SafeInvoke(p => p.Text = "###");
                        materialLabel4.SafeInvoke(p => p.Text = "###");
                        materialLabel5.SafeInvoke(p => p.Text = "000.00");

                        easyDAClient1.WriteItemValue(machineName, serverName, "FOSS.ProFoss.Controller.StartMeasuring", 0);

                        if (plc.finsMemoryAreadWrite(MemoryArea.DM, 18488, 0, 2, arr0))
                            ErrorCounter = 0;
                        else
                        {
                            ErrorCounter++;
                            logger.Warn("Ошибка записи в Omron PLC - {0}", plc.LastError);
                        }

                        if (plc.finsMemoryAreadWrite(MemoryArea.DM, 18486, 0, 2, arr0))
                            ErrorCounter = 0;
                        else
                        {
                            ErrorCounter++;
                            logger.Warn("Ошибка записи в Omron PLC - {0}", plc.LastError);
                        }
                    }
                }

            }
            catch (OpcException ex)
            {
                logger.Error("[{0}] Ошибка в потоке обмена - {1}", ex.LineNumber(), ex.InnerException.Message);
                StopedWithError = true;
            }
            catch (Exception ex)
            {
                logger.Error("[{0}] Ошибка в потоке обмена - {1}", ex.LineNumber(), ex.Message);
                StopedWithError = true;
            }
            finally
            {
                if (fs != null)
                    fs.Dispose();
                if (f != null)
                    f.Dispose();
                if (nirDB != null)
                    nirDB.Dispose();
            }
        }

        private void BackgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            if (!StopedWithError)
            {
                IsRestarting = true;
                Application.Exit();
            }
        }

        private void DataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.ColumnIndex == 3 || e.ColumnIndex == 8)
            {
                e.ThrowException = false;
            }
            else
                logger.Error("[{0}, {1}] {2}", e.ColumnIndex, e.RowIndex, e.Exception.Message);
        }

        private void DataGridView1_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            DataGridView dg = sender as DataGridView;
            if (memoryCache != null)
            {
                try
                {
                    if (e.ColumnIndex == 1 || e.ColumnIndex == 2)
                    {
                        e.Value = memoryCache.RetrieveElement(e.RowIndex, dg.Columns[e.ColumnIndex].DataPropertyName).ToString().PadLeft(2, '0');
                    }
                    else if (e.ColumnIndex == 3)
                    {
                        e.Value = calibrs
                            .Where(p => p.ID == (int)memoryCache.RetrieveElement(e.RowIndex, dg.Columns[e.ColumnIndex].DataPropertyName))
                            .Select(p => p.Name)
                            .FirstOrDefault();
                    }
                    else if (e.ColumnIndex == 8)
                    {
                        e.Value = operators
                            .Where(p => p.ID == (int)memoryCache.RetrieveElement(e.RowIndex, dg.Columns[e.ColumnIndex].DataPropertyName))
                            .Select(p => p.OperatorColumn)
                            .FirstOrDefault();
                    }
                    else
                    {
                        e.Value = memoryCache.RetrieveElement(e.RowIndex, dg.Columns[e.ColumnIndex].DataPropertyName);
                    }
                }
                catch (Exception ex)
                {
                    logger.Error("[{0}] {1}", ex.LineNumber(), ex.Message);
                }
            }
        }

        private void DataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is ComboBox)
            {
                ComboBox comboBox = e.Control as ComboBox;
                comboBox.DropDownClosed -= ComboBox_DropDownClosed;
                comboBox.DropDownClosed += ComboBox_DropDownClosed;
            }
            else if (e.Control is NumericUpDown)
            {
                NumericUpDown numericUpDown = e.Control as NumericUpDown;
                numericUpDown.Leave -= NumericUpDown_Leave;
                numericUpDown.Leave += NumericUpDown_Leave;
            }
        }

        private void NumericUpDown_Leave(object sender, EventArgs e)
        {
            try
            {
                NumericUpDown numericUpDown = sender as NumericUpDown;
                int id = Convert.ToInt32(dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value);
                using (NirDBDB nir = new NirDBDB())
                {
                    switch (dataGridView1.CurrentCell.ColumnIndex)
                    {
                        case 4:
                            nir.NirParams.Where(p => p.ID == id).Set(p => p.StartPercent, (double)numericUpDown.Value).Update();
                            memoryCache.Update();
                            break;
                        case 5:
                            nir.NirParams.Where(p => p.ID == id).Set(p => p.MinPercent, (double)numericUpDown.Value).Update();
                            memoryCache.Update();
                            break;
                        case 6:
                            nir.NirParams.Where(p => p.ID == id).Set(p => p.MaxPercent, (double)numericUpDown.Value).Update();
                            memoryCache.Update();
                            break;
                        case 7:
                            nir.NirParams.Where(p => p.ID == id).Set(p => p.TargetPercent, (double)numericUpDown.Value).Update();
                            memoryCache.Update();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error("{[0}] {1}", ex.LineNumber(), ex.Message);
            }
        }

        private void ComboBox_DropDownClosed(Object sender, EventArgs e)
        {
            try
            {
                ComboBox comboBox = sender as ComboBox;
                int id = Convert.ToInt32(dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value);

                using (NirDBDB nir = new NirDBDB())
                {
                    switch (dataGridView1.CurrentCell.ColumnIndex)
                    {
                        case 3:
                            nir.NirParams.Where(p => p.ID == id).Set(p => p.Calibration, ((Calibr)comboBox.SelectedItem).ID).Update();
                            memoryCache.Update();
                            break;
                        case 8:
                            nir.NirParams.Where(p => p.ID == id).Set(p => p.Operator, ((Operator)comboBox.SelectedItem).ID).Update();
                            memoryCache.Update();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error("[{0}] {1}", ex.LineNumber(), ex.Message);
            }
        }

        private void MaterialTabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (materialTabControl1.SelectedIndex)
            {
                //Вкладка журнал
                case 1:
                    MaterialFlatButton2_Click(materialFlatButton2, null);
                    break;
                //Вкладка отчеты
                case 2:
                    using (NirDBDB nir = new NirDBDB())
                    {
                        comboBox2.DataSource = null;
                        //comboBox2.DataSource = nir.Journals.OrderBy(p => p.DateTime).Select(p => p.RecipeCode).Distinct().ToList();
                    }
                    dataGridView3.Rows.Clear();
                    break;
                //вкладка НАСТРОЙКИ
                case 3:
                    numericUpDown1.Value = NativeMetods.ReadINIInt("General", "Timeout");
                    ipAddressControl1.Text = NativeMetods.ReadINIString("PLC", "IP");
                    numericUpDown2.Value = NativeMetods.ReadINIInt("PLC", "Port");
                    numericUpDown3.Value = NativeMetods.ReadINIInt("PLC", "Timeout");
                    textBox1.Text = NativeMetods.ReadINIString("SQL", "SQLConnectionString");
                    textBox2.Text = NativeMetods.ReadINIString("OPC", "MachineName");
                    textBox3.Text = NativeMetods.ReadINIString("OPC", "Server");
                    textBox4.Text = NativeMetods.ReadINIString("ArchiveVspFile", "LastPosition");
                    textBox5.Text = NativeMetods.ReadINIString("ArchiveVspFile", "Path");
                    comboBox1.DataSource = operators;
                    textBox6.Text = NativeMetods.ReadINIString("SDS", "Availble");
                    break;
            }
        }

        private void DataGridView2_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            DataGridView dg = sender as DataGridView;
            if (memoryCache != null)
            {
                try
                {
                    switch (e.ColumnIndex)
                    {
                        case 2:
                            e.Value = operators
                                .Where(p => p.ID == (int)memoryCache.RetrieveElement(e.RowIndex, dg.Columns[e.ColumnIndex].DataPropertyName))
                                .Select(p => p.OperatorColumn)
                                .FirstOrDefault();
                            break;
                        case 7:
                            e.Value = calibrs
                                .Where(p => p.ID == (int)memoryCache.RetrieveElement(e.RowIndex, dg.Columns[e.ColumnIndex].DataPropertyName))
                                .Select(p => p.Name)
                                .FirstOrDefault();
                            break;
                        default:
                            e.Value = memoryCacheJournal.RetrieveElement(e.RowIndex, dg.Columns[e.ColumnIndex].DataPropertyName);
                            break;
                    }
                }
                catch (Exception ex)
                {
                    logger.Error("[{0}] {1}", ex.LineNumber(), ex.Message);
                }
            }
        }

        //Кнопка ПРИМЕНИТЬ фильтр, вкладка журнал
        private void MaterialFlatButton2_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView2.Rows.Clear();
                if (materialCheckBox2.Checked)
                    retrieverJournal.Filter = " [DateTime] >= '" + dateTimePicker3.Value.ToString("dd.MM.yyyy HH:mm:ss")
                        + "' AND [DateTime] < '" + dateTimePicker4.Value.ToString("dd.MM.yyyy HH:mm:ss") + "' AND ";
                else
                    retrieverJournal.Filter = string.Empty;

                int rowCount = retrieverJournal.RowCount;
                memoryCacheJournal = new Cache(retrieverJournal, 30);
                dataGridView2.SafeInvoke(p => { p.RowCount = rowCount; });
            }
            catch (Exception ex)
            {
                logger.Error("[{0}] {1}", ex.LineNumber(), ex.Message);
            }
        }

        //Обнулить позицию в файле
        private void MaterialFlatButton7_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Внимание! При обнулении будет производиться перечитка файла архива VSP5000 сначала. Продолжить?",
                Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                textBox4.Text = "0";
        }

        //Резервная копия базы данных
        private async void MaterialFlatButton9_Click(object sender, EventArgs e)
        {
            await System.Threading.Tasks.Task.Run(() =>
            {
                SqlConnection connection = null;
                SqlCommand command = null;
                try
                {
                    string connString = NativeMetods.ReadINIString("SQL", "SQLConnectionString");
                    int index1 = connString.IndexOf("Database=") + "Database=".Length;
                    int index2 = connString.IndexOf(';', index1);
                    string connectionstring = connString.Remove(index1, index2 - index1).Replace("Database=", "Database=master");
                    connection = new SqlConnection(connectionstring);
                    connection.Open();
                    string path = System.IO.Path.Combine(Application.StartupPath, "Archive", "NirDB_" + DateTime.Now.ToString("dd_MM_yy HH_mm") + ".bak");
                    command = new SqlCommand(string.Format("BACKUP DATABASE [NirDB] TO DISK = '{0}';", path), connection);
                    command.CommandTimeout = 600;
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    logger.Error("[{0}] {1}", ex.LineNumber(), ex.Message);
                }
                finally
                {
                    if (command != null)
                        command.Dispose();
                    if (connection != null)
                        connection.Dispose();
                }
            });
        }

        //Путь к файлу архива
        private void MaterialFlatButton8_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox5.Text = openFileDialog1.FileName;
            }
        }

        //Сохранить настройки
        private void MaterialFlatButton5_Click(object sender, EventArgs e)
        {
            NativeMetods.WriteINI("General", "Timeout", (int)numericUpDown1.Value);
            NativeMetods.WriteINI("PLC", "IP", ipAddressControl1.Text);
            NativeMetods.WriteINI("PLC", "Port", (int)numericUpDown2.Value);
            NativeMetods.WriteINI("PLC", "Timeout", (int)numericUpDown3.Value);
            NativeMetods.WriteINI("SQL", "SQLConnectionString", textBox1.Text);
            NativeMetods.WriteINI("OPC", "MachineName", textBox2.Text);
            NativeMetods.WriteINI("OPC", "Server", textBox3.Text);
            NativeMetods.WriteINI("ArchiveVspFile", "LastPosition", textBox4.Text);
            NativeMetods.WriteINI("ArchiveVspFile", "Path", textBox5.Text);
            NativeMetods.WriteINI("SDS", "Availble", textBox6.Text);

            MessageBox.Show("Выполнено успешно", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //Очистить базу данных
        private async void MaterialFlatButton4_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Очистить базу данных до даты " +
                dateTimePicker5.Value.ToString("dd.MM.yyyyy HH:mm") + "?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                await System.Threading.Tasks.Task.Run(() =>
                {
                    NirDBDB nir = null;
                    try
                    {
                        nir = new NirDBDB();
                        nir.Command.CommandTimeout = 1000;
                        DateTime date = dateTimePicker5.Value.SetSecond(0).SetMillisecond(0);
                        nir.Journals.Where(p => p.DateTime <= date).Delete();
                        //nir.NirParams.Where(p => p.DateTime <= date).Delete();
                        logger.Info("Очистка журналов весов до даты {0} произведена.", date);
                        logger.Info("Сжимаю базу данных.");
                        nir.Execute("DBCC SHRINKDATABASE(NirDB, 0)");
                        logger.Info("Сжатие выполнено.");
                    }
                    catch (Exception ex)
                    {
                        logger.Error("[{0}] {1}", ex.LineNumber(), ex.Message);
                    }
                    finally
                    {
                        if (nir != null)
                            nir.Dispose();
                    }
                });
            }
        }

        //Добавить оператора
        private void MaterialFlatButton6_Click(object sender, EventArgs e)
        {
            NirDBDB nir = null;
            try
            {
                nir = new NirDBDB();
                if (!nir.Operators.Any(p => p.OperatorColumn == comboBox1.Text))
                {
                    nir.Operators.Value(p => p.OperatorColumn, comboBox1.Text).Insert();
                    operators = nir.Operators.ToList();
                    comboBox1.DataSource = null;
                    comboBox1.DataSource = operators;
                    comboBox1.ValueMember = "ID";
                    comboBox1.DisplayMember = "OperatorColumn";
                    dataGridView1.SafeInvoke(p => (p.Columns[8] as DataGridViewComboBoxColumn).DataSource = null);
                    dataGridView1.SafeInvoke(p => (p.Columns[8] as DataGridViewComboBoxColumn).DataSource = operators);
                    dataGridView1.SafeInvoke(p => (p.Columns[8] as DataGridViewComboBoxColumn).ValueMember = "ID");
                    dataGridView1.SafeInvoke(p => (p.Columns[8] as DataGridViewComboBoxColumn).DisplayMember = "OperatorColumn");
                }
            }
            catch (Exception ex)
            {
                logger.Error("[{0}] {1}", ex.LineNumber(), ex.Message);
            }
            finally
            {
                if (nir != null)
                    nir.Dispose();
            }
        }

        //Удалить оператора
        private void MaterialFlatButton3_Click(object sender, EventArgs e)
        {
            NirDBDB nir = null;
            try
            {
                nir = new NirDBDB();
                nir.Operators.Where(p => p.OperatorColumn == comboBox1.Text).Delete();
                operators = nir.Operators.ToList();
                comboBox1.DataSource = null;
                comboBox1.DataSource = operators;
                comboBox1.ValueMember = "ID";
                comboBox1.DisplayMember = "OperatorColumn";
                dataGridView1.SafeInvoke(p => (p.Columns[8] as DataGridViewComboBoxColumn).DataSource = null);
                dataGridView1.SafeInvoke(p => (p.Columns[8] as DataGridViewComboBoxColumn).DataSource = operators);
                dataGridView1.SafeInvoke(p => (p.Columns[8] as DataGridViewComboBoxColumn).ValueMember = "ID");
                dataGridView1.SafeInvoke(p => (p.Columns[8] as DataGridViewComboBoxColumn).DisplayMember = "OperatorColumn");
            }
            catch (Exception ex)
            {
                logger.Error("[{0}] {1}", ex.LineNumber(), ex.Message);
            }
            finally
            {
                if (nir != null)
                    nir.Dispose();
            }
        }

        private TotalPartOfString GetTotalPartOfString(string mes)
        {
            TotalPartOfString totalPartOfString = new TotalPartOfString();
            mes = mes.Replace("\0", "").Trim();
            string patternLoad = @"\A(\d+)\s+(\d{2}[.]\d{2}[.]\d{2}\s+\d{2}[:]\d{2}[:]\d{2}"
                + @"\.*\d{0,2})\s+(\S+):\s+Задан:\s+\[(\S+)\]\s+\[(\S+)\]\s+\[(\S+)\]\z";
            string patternStart = @"\A(\d+)\s+(\d{2}[.]\d{2}[.]\d{2}\s+\d{2}[:]\d{2}[:]\d{2}"
                + @"\.*\d{0,2})\s+(\S+):\s+Пуск:\s+\[(\S+)\]\s+\[(\S+)\]\s+\[(\S+)\]\z";

            Match matchdate;
            //Задаю регулярные выражения для поиска в строке
            if (Regex.IsMatch(mes, patternLoad, RegexOptions.IgnoreCase))
            {
                matchdate = Regex.Match(mes, patternLoad, RegexOptions.IgnoreCase);
                totalPartOfString.Start = false;
            }
            else
            {
                matchdate = Regex.Match(mes, patternStart, RegexOptions.IgnoreCase);
                totalPartOfString.Start = true;
            }

            string num = matchdate.Groups[1].Value;
            string date = matchdate.Groups[2].Value;
            totalPartOfString.DeviceName = matchdate.Groups[3].Value;
            totalPartOfString.RecipeNum = matchdate.Groups[4].Value;
            totalPartOfString.TaskNum = matchdate.Groups[5].Value;
            totalPartOfString.MixFor = matchdate.Groups[6].Value;
            int.TryParse(num, out totalPartOfString.Num);
            DateTime.TryParse(date, out totalPartOfString.DateT);
            return totalPartOfString;
        }

        private int old_rowCount = 0;
        private int old_rowCountJournal = 0;

        //Обновить таблицу журнала
        private void RefreshJournal()
        {
            //Вызываем перечитку журнала
            try
            {
                int rowCountJournal = retrieverJournal.RowCount;
                if (old_rowCountJournal != rowCountJournal)
                {
                    memoryCacheJournal = new Cache(retrieverJournal, 30);
                    dataGridView2.SafeInvoke(p => { p.RowCount = rowCountJournal; });
                    old_rowCountJournal = rowCountJournal;
                }
            }
            catch { }
        }

        //Обновить таблицу параметров
        private void RefreshParams()
        {
            //Вызываем перечитку таблицы параметров
            try
            {
                int rowCount = retriever.RowCount;
                if (old_rowCount != rowCount)
                {
                    memoryCache = new Cache(retriever, 30);
                    dataGridView1.SafeInvoke(p => { p.RowCount = rowCount; });
                    old_rowCount = rowCount;
                }
            }
            catch { }
        }

        //Сформировать отчет
        private async void MaterialFlatButton10_Click(object sender, EventArgs e)
        {
            dataGridView3.Rows.Clear();
            await System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    using (NirDBDB nir = new NirDBDB())
                    {
                        string recipe = comboBox2.SafeInvoke(p => p.Text);
                        //List<Journal> journal = nir.Journals.Where(p => p.RecipeCode == recipe)
                        //    .LoadWith(p => p.FKOperator).LoadWith(p => p.Calibr).OrderBy(p => p.DateTime).ToList();
                        //foreach (Journal line in journal)
                        //{
                        //    dataGridView3.SafeInvoke(p => p.Rows.Add(line.DateTime, line.FKOperator.OperatorColumn,
                        //        line.WaterPercent, line.MoiPercent, line.NumberOtves, line.Calibr.Name));
                        //}
                    }
                }
                catch (Exception ex)
                {
                    logger.Error("[{0}] {1}", ex.LineNumber(), ex.Message);
                }
            });
        }

        //Вывести в Exel
        private async void MaterialFlatButton11_Click(object sender, EventArgs e)
        {
            if (dataGridView3.Rows.Count.Equals(0))
                return;

            string reportPath = Path.Combine(Application.StartupPath, "templates", "Report.xlsx");
            string TemplatePath = Path.Combine(Application.StartupPath, "templates", "Report_template.xlsx");

            await System.Threading.Tasks.Task.Run(() =>
            {
                using (XLWorkbook excelWorkbook = new XLWorkbook(TemplatePath))
                {
                    try
                    {
                        IXLWorksheet ws = excelWorkbook.Worksheet(1);

                        var firstCell = ws.FirstCellUsed();
                        var lastCell = ws.LastCellUsed();

                        IXLRange range = ws.Range(firstCell.Address, lastCell.Address);

                        IXLCell foundCell = null;

                        #region Заполнение дополнительных полей на листе
                        //Это надо делать сначала т.к. потом разметка сдвинется

                        //Дата начала
                        foundCell = range.Find("{datefrom}");
                        if (foundCell != null)
                            foundCell.Value = dataGridView3.SafeInvoke(p => p[0, 0].Value);

                        //Дата окончания
                        int lastRow = dataGridView3.Rows.Count - 1;
                        foundCell = range.Find("{dateto}");
                        if (foundCell != null)
                            foundCell.Value = dataGridView3.SafeInvoke(p => p[0, lastRow].Value);

                        //Номер рецепта
                        foundCell = range.CellsUsed().AsParallel().Where(p => p.Value.ToString().Contains("{recipenum}")).FirstOrDefault();
                        if (foundCell != null)
                            foundCell.Value = foundCell.Value.ToString().Replace("{recipenum}", comboBox2.SafeInvoke(p => p.GetItemText(comboBox2.SelectedItem)));
                        #endregion

                        //Начало таблицы
                        DataGridToExel dataGridToExel = new DataGridToExel();
                        dataGridToExel.ConvertDataGridViewToExelWithFormatting(ref ws, dataGridView3);
                        //Параметры печати
                        ws.PageSetup.CenterHorizontally = true;

                        ws.SheetView.View = XLSheetViewOptions.PageLayout;
                        excelWorkbook.SaveAs(reportPath);
                        Process.Start(new ProcessStartInfo(reportPath) { UseShellExecute = true });
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show("Ошибка открытия файла отчета в Excel. Закройте все окна Excel и повторите формирование отчета.",
                            Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        logger.Error("[{0}] {1}", ex.LineNumber(), ex.Message);
                    }
                    catch (Exception ex)
                    {
                        logger.Error("[{0}] {1}", ex.LineNumber(), ex.Message);
                    }
                }
            });
        }

        //Обновление checkBoxColumn
        private void DataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex == dataGridView1.Columns["Column23"].Index)
            {
                dataGridView1.EndEdit();
            }
        }

        private void DataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dataGridView1.Columns["Column23"].Index)
            {
                try
                {
                    DataGridViewCheckBoxCell checkBox = dataGridView1[e.ColumnIndex, e.RowIndex] as DataGridViewCheckBoxCell;
                    int id = Convert.ToInt32(dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value);
                    using (NirDBDB nir = new NirDBDB())
                    {
                        nir.NirParams.Set(p => p.EnableCorrection, Convert.ToBoolean(checkBox.EditingCellFormattedValue)).Update();
                    }
                    memoryCache.Update();
                }
                catch (Exception ex)
                {
                    logger.Error("{[0}] {1}", ex.LineNumber(), ex.Message);
                }
            }
        }
    }
}
