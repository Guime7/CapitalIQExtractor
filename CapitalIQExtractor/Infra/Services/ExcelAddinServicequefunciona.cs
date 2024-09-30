// using System;
// using System.Diagnostics;
// using System.Runtime.InteropServices;
// using System.Text;
// using CapitalIQExtractor.Core.Interfaces;
// using Excel = Microsoft.Office.Interop.Excel;
//
// namespace CapitalIQExtractor.Infra.Services;
//
// public class ExcelAddinService : IExcelAddinService
// {
//     private const string EXCEL_CLASS_NAME = "EXCEL7";
//     private const uint DW_OBJECTID = 0xFFFFFFF0;
//     private static Guid rrid = new Guid("{00020400-0000-0000-C000-000000000046}");
//
//     public delegate bool EnumChildCallback(int hwnd, ref int lParam);
//
//     [DllImport("Oleacc.dll")]
//     public static extern int AccessibleObjectFromWindow(int hwnd, uint dwObjectID, byte[] riid,
//         ref Excel.Window ptr);
//
//     [DllImport("User32.dll")]
//     public static extern bool EnumChildWindows(int hWndParent, EnumChildCallback lpEnumFunc, ref int lParam);
//
//     [DllImport("User32.dll")]
//     public static extern int GetClassName(int hWnd, StringBuilder lpClassName, int nMaxCount);
//
//     private static Excel.Application GetExcelInterop(int? processId = null)
//     {
//         var p = processId.HasValue
//             ? Process.GetProcessById(processId.Value)
//             : Process.Start("C:\\Program Files\\Microsoft Office\\root\\Office16\\excel.exe", "/x");
//         try
//         {
//             return new ExcelAddinService().SearchExcelInterop(p);
//         }
//         catch (Exception)
//         {
//             Debug.Assert(p != null, "p != null");
//             return GetExcelInterop(p.Id);
//         }
//     }
//
//     private bool EnumChildFunc(int hwndChild, ref int lParam)
//     {
//         var buf = new StringBuilder(128);
//         GetClassName(hwndChild, buf, 128);
//         if (buf.ToString() == EXCEL_CLASS_NAME)
//         {
//             lParam = hwndChild;
//             return false;
//         }
//
//         return true;
//     }
//
//     private Excel.Application SearchExcelInterop(Process p)
//     {
//         Excel.Window ptr = null;
//         int hwnd = 0;
//
//         int hWndParent = (int)p.MainWindowHandle;
//         EnumChildWindows(hWndParent, EnumChildFunc, ref hwnd);
//         int hr = AccessibleObjectFromWindow(hwnd, DW_OBJECTID, rrid.ToByteArray(), ref ptr);
//         return ptr.Application;
//     }
//
//     public async Task<Excel.Application> GetExcelInstance()
//     {
//         var excel = await Task.Run(() =>
//             {
//                 var instance = GetExcelInterop();
//                 instance.Visible = false;
//                 return instance;
//         });
//
//         //disable alters and other to make excel faster
//         excel.DisplayAlerts = false;
//         excel.ScreenUpdating = false;
//         excel.EnableEvents = false;
//         // excel.Calculation = Excel.XlCalculation.xlCalculationManual;
//         return excel;
//     }
// }