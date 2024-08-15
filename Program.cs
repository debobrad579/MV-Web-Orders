using System;
using System.Threading;
using System.Threading.Tasks;

using FlaUI.Core;
using FlaUI.UIA3;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Input;

class Program
{
    static void Main()
    {
        Console.Write("Enter a order number: ");
        string? orderNumber = Console.ReadLine();

        while (!CheckOrderExists(orderNumber))
        {
            Console.Write("Invalid order number. Enter a valid order number: ");
            orderNumber = Console.ReadLine();
        }

        using (var automation = new UIA3Automation())
        {
            Application app = Application.Launch(@"c:\Program Files\Microvellum\Toolbox OEM 2022\toolbox.exe");
            ConditionFactory cf = new ConditionFactory(new UIA3PropertyLibrary());
            Thread.Sleep(500);
            var window = app.GetMainWindow(automation);

            while (true)
            {
                window = app.GetMainWindow(automation);
                var continueBtn = window.FindFirstDescendant(cf.ByAutomationId("btnStartApplication")).AsButton();

                if (continueBtn != null)
                {
                    continueBtn.Invoke();
                    break;
                }

                Thread.Sleep(500);
            }
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
