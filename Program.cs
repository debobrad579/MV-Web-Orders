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

class Program
{
    static void Main()
    {
        // Console.Write("Enter a order number: ");
        // string? orderNumber = Console.ReadLine();

        // while (!CheckOrderExists(orderNumber))
        // {
        //     Console.Write("Invalid order number. Enter a valid order number: ");
        //     orderNumber = Console.ReadLine();
        // }

        UIA3Automation? automation = new UIA3Automation();
        Application app = Application.Launch(@"c:\Program Files\Microvellum\Toolbox OEM 2022\toolbox.exe");
        ConditionFactory cf = new ConditionFactory(new UIA3PropertyLibrary());

        Window? window = null;
        Button? continueBtn = null;

        while (true)
        {
            window = app.GetMainWindow(automation);
            continueBtn = window.FindFirstDescendant(cf.ByAutomationId("btnStartApplication")).AsButton();

            if (continueBtn != null)
            {
                continueBtn.Invoke();
                break;
            }

            Thread.Sleep(500);
        }

        Menu? menu = null;

        while (true)
        {
            window = app.GetMainWindow(automation);
            menu = window.FindFirstDescendant(cf.ByName("MenuBar1")).AsMenu();

            if (menu != null && menu.BoundingRectangle.Top > 200)
            {
                var menuRect = menu.BoundingRectangle;
                var mousePos = Mouse.Position;
                Mouse.Position = new Point(menuRect.Left + 150, menuRect.Top + 10);
                Mouse.Click();
                Mouse.Position = mousePos;
                break;
            }

            Thread.Sleep(500);
        }

    }

    private static bool CheckOrderExists(string? orderNumber)
    {
        foreach (string file in Directory.EnumerateFiles(@"A:\8 - Website\WEBSITE ORDERS", "*.*", SearchOption.AllDirectories))
        {
            if (Path.GetFileName(file) == $"order-{orderNumber}.xml")
            {
                return true;
            }
        }

        return false;
    }
}
