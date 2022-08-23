using System;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using DataModels;
using LinqToDB;

namespace OPCtoOmron
{
    static class Program
    {
        public static readonly MaterialSkin.MaterialSkinManager materialSkinManager = MaterialSkin.MaterialSkinManager.Instance;
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            if (!Properties.Settings.Default.IsRestarting && SingleInstance.SingleApplication.IsAlreadyRunning())
            {
                //Если приложение уже запущено
                SingleInstance.SingleApplication.SwitchToCurrentInstance();
            }
            else
            {
                NirDBDB nir = null;
                try
                {
                    nir = new NirDBDB();
                    var sp = nir.DataProvider.GetSchemaProvider();
                    var dbSchema = sp.GetSchema(nir);
                    if (!dbSchema.Tables.Any(t => t.TableName == "Calibr"))
                    {
                        nir.CreateTable<Calibr>();
                        nir.Calibrs
                            .Value(p => p.Code, 7000)
                            .Value(p => p.Name, "Dairy-Klin")
                            .Insert();
                        nir.Calibrs
                            .Value(p => p.Code, 2000)
                            .Value(p => p.Name, "Poultry-Klin")
                            .Insert();
                        nir.Calibrs
                            .Value(p => p.Code, 1000)
                            .Value(p => p.Name, "Swine-Klin")
                            .Insert();
                    }

                    if (!dbSchema.Tables.Any(t => t.TableName == "Journal"))
                    {
                        nir.CreateTable<Journal>();
                    }

                    if (!dbSchema.Tables.Any(t => t.TableName == "NIR_Params"))
                    {
                        nir.CreateTable<NirParam>();
                    }

                    if (!dbSchema.Tables.Any(t => t.TableName == "Operators"))
                    {
                        nir.CreateTable<Operator>();
                        nir.Operators.Value(p => p.OperatorColumn, "Петров").Insert();
                    }

                    nir.Dispose();
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    materialSkinManager.Theme = MaterialSkin.MaterialSkinManager.Themes.DARK;
                    Application.Run(new MainForm());
                }
                catch (Exception ex)
                {
                    if (nir != null)
                        nir.Dispose();
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
