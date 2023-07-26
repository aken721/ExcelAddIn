using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{
    public partial class ThisAddIn
    {
        public static Excel.Application app;         //声明一个excel变量
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            app = Globals.ThisAddIn.Application;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
