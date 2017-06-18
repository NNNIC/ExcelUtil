#define WINAPP

using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;

#if WINAPP
using System.Windows.Forms;
#endif

using _Application = Microsoft.Office.Interop.Excel.Application;

/*
    Excel Utility (C) 2017 NNNIC /  ZIP LICENSE

    Futures

    #1. Use Microsoft soft original dll only
    #2. Quick Access
    #3. Capable to attach to the opened excel window
    #4. Simple Interface

    <<Sammary>>

    BookCtr   book = ExcelUtil.OpenBook(path);
    SheetCtr  sheet = app.GetSheet(sheetname);

    object[,] values =  sheet.GetValues(numofrow=null,numofcol=null); //default : Used Range
    sheet.SetValues(values); 

    book.Close(); 
    book.Write();
    book.WriteAs(path)

    BookCtr book = book.OpenBookInApp(path);  //Open a book ont the excel window

    BookCtr book = ExcelUtil.AttachBook(path); //Attach to a existing excelwindow


*/

public partial class ExcelUtil
{
    static void ShowError(string msg)
    {
#if WINAPP
        MessageBox.Show(msg);
#else
        System.Diagnostics.Debug.WriteLine(msg);
#endif
    }

    public class AppCtr
    {
        public _Application m_application { get; private set;}
        public Workbooks    m_workbooks { get; private set; }
        public int          m_refcnt;

        public AppCtr() { }
        public AppCtr(_Application app, Workbooks wbs)
        {
            m_application = app;
            m_workbooks   = wbs;
        }

        public _Application New() {
            if (m_application==null)
            {
                m_application = new _Application();
                m_workbooks   = m_application.Workbooks;
            }
            m_refcnt++;
            return m_application;
        }

        public void Delete()
        {
            m_refcnt--;
            if (m_refcnt<=0)
            {
                _dispose();
            }
        }

        private void _dispose()
        {
            if (m_workbooks!=null)
            {
                try {
                    Marshal.ReleaseComObject(m_workbooks);  
                } finally
                {
                    m_workbooks = null;
                }
            }

            if (m_application!=null)
            {
                try {
                    Marshal.ReleaseComObject(m_application);  
                } finally
                {
                    m_application = null;
                }
            }
        }

        ~AppCtr()
        {
            _dispose();
        }
    }

    public class BookCtr : IDisposable
    {
               AppCtr         m_appCtr;
               _Application   m_app { get { return m_appCtr.m_application; } }
               Workbooks      m_wbs { get { return m_appCtr.m_workbooks;   } }
        public string         m_path;
        public Workbook       m_wb;
               List<SheetCtr> m_sheetCtr_list = new List<SheetCtr>();

        public BookCtr() { }
        public BookCtr(_Application app, Workbooks wbs, Workbook wb, string path)
        {
            m_appCtr = new AppCtr(app, wbs);
            m_wb     = wb;
            m_path   = path;
        }
        
        public BookCtr OpenNewBook(string path)
        {
            var wb = m_wbs.Open(path);
            var newbookctr = new BookCtr();
            m_appCtr.New();
            newbookctr.m_wb = wb;
            return newbookctr;
        }
        
        public void SetVisible(bool b)
        {
            m_app.Visible = b;
        }     

        public SheetCtr GetSheet(string sheetname) {
            return _getsheet(sheetname);
        }
        public SheetCtr GetActiveSheet() {
            Worksheet worksheet = null;
            string    name = null;
            try { worksheet = (Worksheet)m_wb.ActiveSheet; name = worksheet.Name; } catch { worksheet=null; return null; } finally {  Marshal.ReleaseComObject(worksheet); }
            return   _getsheet(name);
        }
        public SheetCtr GetSheet(int index_base0) {
            Worksheet worksheet = null;
            string name = null;
            try {
                worksheet = m_wb.Sheets[index_base0+1]; name = worksheet.Name;
            } catch { worksheet = null; return null; } finally { Marshal.ReleaseComObject(worksheet); }            
            return _getsheet(name);
        }
        private SheetCtr _getsheet(string name)
        {
            var find = m_sheetCtr_list.Find(s=>s.m_sheet.Name==name);
            if (find!=null)
            {
                return find;
            }
            Worksheet worksheet = null;
            SheetCtr  sheetctr = null;
            try {
                worksheet = m_wb.Sheets[name];
                sheetctr  = new SheetCtr(m_appCtr,this,worksheet);
                m_sheetCtr_list.Add(sheetctr);
            }
            catch
            {
                if (worksheet!=null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
                return null;
            }
            return sheetctr;
        }

        public void Write()
        {
            bool bOk = false;
            if (m_wb!=null)
            {
                var bSame = false;
                try
                {
                    bSame = Path.GetFullPath(m_wb.Path)==m_path;
                }
                catch { }
                if (bSame)
                {
                    try {
                        m_wb.Save();  
                        bOk = true;
                    } catch { }
                }
                else
                {
                    try {
                        m_wb.SaveAs(m_path);
                        bOk = true;
                    } catch { }
                }
            }
            if (!bOk)
            {
                ShowError("Faild to write");
            }
        }
        public void WriteAs(string path) {
            bool bOk = false;
            if (m_wb!=null)
            {
                try {
                    m_wb.SaveAs(path);
                    bOk = true;
                } catch { }
            }
            if (!bOk)
            {
                ShowError("Failed to write");
            }
        }
        public void Close() {
            Dispose();
        }

#region Dispose

        public void Dispose() //ref http://ufcpp.net/study/csharp/rm_disposable.html
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        } 

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // managed 
                m_sheetCtr_list.ForEach(s=>s.Dispose());
                m_appCtr.Delete();
            }
            try {
                if (m_wb!=null) {
                    try { m_wb.Close(false); } catch { }
                    Marshal.ReleaseComObject(m_wb);
                    m_wb = null;
                }
            } catch {  m_wb = null; }
        }

        ~BookCtr()
        {
            Dispose(false);
        }
#endregion
    }
    public class SheetCtr: IDisposable
    {
               AppCtr       m_appctr;
               BookCtr      m_bookctr;
               _Application m_app      { get { return m_appctr.m_application;} }
               Workbook     m_workbook { get { return m_bookctr.m_wb;        } }
        public Worksheet    m_sheet    { get; private set;                     }

        public SheetCtr() { }
        public SheetCtr(AppCtr appctr, BookCtr bookctr, Worksheet sheet)
        {
            m_appctr  = appctr;
            m_bookctr = bookctr;
            m_sheet   = sheet;
        }

        public object[,] GetValues(object numofrows=null,object numofcols=null)
        {
            Range range_values = null;
            object[,] values = null;
            try {
                if (numofcols==null && numofrows==null)
                {
                    range_values = m_sheet.UsedRange;
                    values = range_values.Value2;
                }
                else 
                {
                    string colstr = null;
                    if (numofcols is string)
                    {
                        colstr = numofcols.ToString();
                    }
                    else if (numofcols is int)
                    {
                        colstr = _toAlphabet_nb1((int)numofcols);
                    }
                    else
                    {
                        throw new SystemException();
                    }
                    colstr += numofrows.ToString();
                    range_values = m_sheet.Range["A1",colstr];
                }
                values = range_values.Value2;
            }finally
            {
                Marshal.ReleaseComObject(range_values);
            }
            return values;
        }
        public void SetValues(object[,] values,bool bUseSetting=true)
        {
            Range range_values = null;
            try {
                var len2 = values.GetLength(0);
                var len1 = values.GetLength(1);

                var colstr = _toAlphabet_nb1(len1) + (len2).ToString();
                range_values = m_sheet.Range["A1",colstr];
                if (bUseSetting)
                {
                    //range_values.Font.Size = 30;
                    range_values.NumberFormatLocal="@";
                }
                range_values.Value2 = values;
            } finally
            {
                Marshal.ReleaseComObject(range_values);
            }            
        }

#region Dispose
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
            }
            try {
                if (m_sheet!=null)
                {
                    Marshal.ReleaseComObject(m_sheet);
                }
            } finally { m_sheet = null; }
        }
#endregion
    }

    public static BookCtr OpenBook(string path)
    {
        var fullpath = Path.GetFullPath(path);

        _Application app = null;
        Workbooks    wbs = null;
        Workbook     wb  = null;

        try {
            app = new _Application();
            app.DisplayAlerts = false;
            wbs = app.Workbooks;
            if (File.Exists(fullpath))
            {
                wb = wbs.Open(path);
            }
            else
            {
                wb = wbs.Add();
            }
        }
        catch
        {
            if (wb!=null)
            {   
                try {  wb.Close(false); } catch { }
                Marshal.ReleaseComObject(wb);
                wb = null;
            }
            if (wbs!=null)
            {
                Marshal.ReleaseComObject(wbs);
                wbs = null;
            }
            if (app!=null)
            {
                try { app.Quit(); } catch { }
                Marshal.ReleaseComObject(app);
                app = null;
            }
            return null;
        }
        return new BookCtr(app,wbs,wb,fullpath);   
    }
    public static BookCtr AttackBook(string path)
    {
        var fullpath = Path.GetFullPath(path);

        _Application app = null;
        Workbooks    wbs = null;
        Workbook     wb  = null;
        
        try { app = (_Application)Interaction.GetObject(fullpath,"Excel.Application"); } catch {app=null;}
        if (app==null) return null;

        try {
            app.DisplayAlerts = false;
            wbs = app.Workbooks;
            if (File.Exists(fullpath))
            {
                wb = wbs.Open(path);
            }
            else
            {
                wb = wbs.Add();
                
                
            }
        }
        catch
        {
            if (wb!=null)
            {   
                try {  wb.Close(false); } catch { }
                Marshal.ReleaseComObject(wb);
                wb = null;
            }
            if (wbs!=null)
            {
                Marshal.ReleaseComObject(wbs);
                wbs = null;
            }
            if (app!=null)
            {
                Marshal.ReleaseComObject(app);
                app = null;
            }
            return null;
        }
        return new BookCtr(app,wbs,wb,fullpath);   
    }


#region tool

    //エクセルカラム(ベース１)をアルファベット文字へ
    // ref https://support.microsoft.com/ja-jp/help/833402/how-to-convert-excel-column-numbers-into-alphabetical-characters
    // ref https://memo-c-sharp.blogspot.jp/2015/09/excel.html
    // ref https://stackoverflow.com/questions/181596/how-to-convert-a-column-number-eg-127-into-an-excel-column-eg-aa
    static string _toAlphabet_nb1(int columnNumberOneBased)
    {
        int baseValue = Convert.ToInt32('A');
        int columnNumberZeroBased = columnNumberOneBased - 1;

        string ret = "";

        if (columnNumberOneBased > 26)
        {
          ret = _toAlphabet_nb1(columnNumberZeroBased / 26) ;
        }

        return ret + Convert.ToChar(baseValue + (columnNumberZeroBased % 26) );
    }

    //エクセルカラム文字列を数字(ベース１)へ
    static int _toNum_nb1(string str)
    {
        if (str.Length>0 &&  char.IsDigit(str[0]))
        {
            int x = 0;
            if (int.TryParse(str,out x))
            {
                return x;
            }
            return 0;
        }
        var columnStr = str.ToUpper();
 
        double result = 0;
        for (int i = 0; i < columnStr.Length; i++)
        {
            int x = columnStr.Length - i - 1;
            int num = Convert.ToChar(columnStr[x]) - 64;
 
            result += num * Math.Pow(26, i);
        }
        return Convert.ToInt32(result);
    }

    static int GetSheetIndex(Workbook wb,string name)
    {
        for(var i = 0; i<wb.Sheets.Count; i++)
        {
            var s = (Worksheet)wb.Sheets[i];
            if (s.Name == name)
            {
                return i;
            }
        }
        return -1;
    } 
#endregion
}