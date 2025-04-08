using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;


namespace RibbonHelloWorld
{
    public class Class1 : ExcelRibbon
    {
        public void OnButtonPressed(IRibbonControl control)
        {
            MessageBox.Show("Hello, World from Excel-DNA!", "Excel-DNA Custom Ribbon");
        }

        public string GetButtonLabel(IRibbonControl control)
        {
            return "Hello World";
        }
    }
    public static class Functions
    {
        [ExcelFunction(Description = "Hello World function")]
        public static string HelloWorld()
        {
            return "Hello, World from Excel-DNA!";
        }
    }
}
