using System;
using System.Threading;
using System.Threading.Tasks;
using FlaUI.Core;
using FlaUI.UIA3;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Input;
using System.CodeDom;
using System.Drawing;
using FlaUI.Core.Tools;
using FlaUI.Core.WindowsAPI;

class Program
{
    static void Main()
    {
        Console.Write("Enter a order number: ");
        string? orderNumber = Console.ReadLine();
        string? orderDirectoryPath = GetOrderPath(orderNumber);

        while (orderDirectoryPath is null)
        {
            Console.Write("Invalid order number. Enter a valid order number: ");
            orderNumber = Console.ReadLine();
        }

        UIA3Automation? automation = new();
        Application app = Application.Launch(@"c:\Program Files\Microvellum\Toolbox OEM 2022\toolbox.exe");
        ConditionFactory cf = new(new UIA3PropertyLibrary());
        Window? window = app.GetMainWindow(automation);

        Retry.WhileException(() =>
        {
            window = app.GetMainWindow(automation);
            Button continueBtn = window.FindFirstDescendant(cf.ByAutomationId("btnStartApplication")).AsButton();
            continueBtn.Invoke();
        }, TimeSpan.FromSeconds(60), null, true);

        Retry.WhileException(() =>
        {
            window = app.GetMainWindow(automation);
            Menu menu = window.FindFirstDescendant(cf.ByName("MenuBar1")).AsMenu();
            Rectangle menuRect = menu.BoundingRectangle;
            if (menuRect.Top < 150) { throw new Exception("Too close to the top."); }
            var mousePos = Mouse.Position;
            Mouse.Click(new Point(menuRect.Left + 150, menuRect.Top + 10));
            Thread.Sleep(50);
            Mouse.Click(new Point(menuRect.Left + 150, menuRect.Top + 285));
            Mouse.Position = mousePos;
        }, TimeSpan.FromSeconds(60), null, true);

        Retry.WhileException(() =>
        {
            window = app.GetMainWindow(automation);
            Button browseFileBtn = window.FindFirstDescendant(cf.ByAutomationId("Button1")).AsButton();
            browseFileBtn.Click();
        }, TimeSpan.FromSeconds(60), null, true);

        window = app.GetMainWindow(automation);
        Keyboard.Type(orderDirectoryPath);
        Keyboard.Press(VirtualKeyShort.ENTER);
        Keyboard.Type($"order-{orderNumber}.xml");
        Keyboard.Press(VirtualKeyShort.ENTER);

        Retry.WhileException(() =>
        {
            var buttonBar = window.FindFirstDescendant(cf.ByAutomationId("ButtonBar1"));
            Rectangle buttonBarRect = buttonBar.BoundingRectangle;
            Mouse.Click(new Point(buttonBarRect.Left + 50, buttonBarRect.Top + 12));
        }, TimeSpan.FromSeconds(60), null, true);

        // if (window.FindFirstDescendant(cf.ByAutomationId("TextBox")) is not null)
        // {
        //     Console.WriteLine("Project already exists.");
        //     Console.ReadKey();
        //     return;
        // }

        Retry.WhileException(() =>
        {
            window = app.GetMainWindow(automation);
            Button okButton = window.FindFirstDescendant(cf.ByName("OK")).AsButton();
            okButton.Invoke();
        }, TimeSpan.FromSeconds(120), null, true);

        Retry.WhileException(() =>
        {
            window = app.GetMainWindow(automation);
            var bbProject = window.FindFirstDescendant(cf.ByAutomationId("bbProject"));
            Rectangle bbProjectRect = bbProject.BoundingRectangle;
            Mouse.Click(new Point(bbProjectRect.Left + 50, bbProjectRect.Top + 12));
        }, TimeSpan.FromSeconds(120), null, true);

        Retry.WhileException(() =>
        {
            window = app.GetMainWindow(automation);
            TextBox searchBar = window.FindFirstDescendant(cf.ByAutomationId("txtSearch")).AsTextBox();
            searchBar.Text = $"Order#{orderNumber}";
        }, TimeSpan.FromSeconds(120), null, true);
    }

    private static string? GetOrderPath(string? orderNumber)
    {
        foreach (string file in Directory.EnumerateFiles(@"A:\8 - Website\WEBSITE ORDERS", "*.*", SearchOption.AllDirectories))
        {
            if (Path.GetFileName(file) == $"order-{orderNumber}.xml")
            {
                return Path.GetDirectoryName(file);
            }
        }

        return null;
    }
}
