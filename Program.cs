using FlaUI.Core;
using FlaUI.UIA3;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Input;
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
            orderDirectoryPath = GetOrderPath(orderNumber);
        }

        UIA3Automation? automation = new();
        Application app = Application.Launch(@"c:\Program Files\Microvellum\Toolbox OEM 2022\toolbox.exe");
        ConditionFactory cf = new(new UIA3PropertyLibrary());
        Window? window = app.GetMainWindow(automation);

        void WhileException(Action<Window> func, int timeoutSeconds = 60)
        {
            Retry.WhileException(() => { Window window = app.GetMainWindow(automation); func(window); }, TimeSpan.FromSeconds(timeoutSeconds), null, true);
        }

        void MouseClick(Point point)
        {
            Point startingMousePosition = Mouse.Position;
            Mouse.Click(point);
            Mouse.Position = startingMousePosition;
        }

        void MouseDoubleClick(Point point)
        {
            Point startingMousePosition = Mouse.Position;
            Mouse.DoubleClick(point);
            Mouse.Position = startingMousePosition;
        }

        WhileException(window =>
        {
            Button continueBtn = window.FindFirstDescendant(cf.ByAutomationId("btnStartApplication")).AsButton();
            continueBtn.Invoke();
        });

        WhileException(window =>
        {
            Menu menu = window.FindFirstDescendant(cf.ByName("MenuBar1")).AsMenu();
            if (menu.BoundingRectangle.Top < 150) { throw new Exception("Too close to the top."); }
            MouseClick(new Point(menu.BoundingRectangle.Left + 150, menu.BoundingRectangle.Top + 10));
            Thread.Sleep(50);
            MouseClick(new Point(menu.BoundingRectangle.Left + 150, menu.BoundingRectangle.Top + 285));
        });

        WhileException(window =>
        {
            Button browseFileBtn = window.FindFirstDescendant(cf.ByAutomationId("Button1")).AsButton();
            MouseClick(new Point(browseFileBtn.BoundingRectangle.Left + 5, browseFileBtn.BoundingRectangle.Top + 5));
        });

        window = app.GetMainWindow(automation);
        Keyboard.Type(orderDirectoryPath);
        Keyboard.Press(VirtualKeyShort.ENTER);
        Keyboard.Type($"order-{orderNumber}.xml");
        Keyboard.Press(VirtualKeyShort.ENTER);

        WhileException(window =>
        {
            AutomationElement buttonBar = window.FindFirstDescendant(cf.ByAutomationId("ButtonBar1"));
            MouseClick(new Point(buttonBar.BoundingRectangle.Left + 50, buttonBar.BoundingRectangle.Top + 12));
        });

        WhileException(window =>
        {
            Button okButton = window.FindFirstDescendant(cf.ByName("OK")).AsButton();
            okButton.Invoke();
        });

        WhileException(window =>
        {
            AutomationElement bbProject = window.FindFirstDescendant(cf.ByAutomationId("bbProject"));
            MouseClick(new Point(bbProject.BoundingRectangle.Left + 50, bbProject.BoundingRectangle.Top + 12));
        });

        WhileException(window =>
        {
            TextBox searchBar = window.FindFirstDescendant(cf.ByAutomationId("txtSearch")).AsTextBox();
            searchBar.Text = $"Order#{orderNumber}";
        });

        Thread.Sleep(500);

        AutomationElement treeView = window.FindFirstDescendant(cf.ByAutomationId("TreeView1"));
        MouseDoubleClick(new Point(treeView.BoundingRectangle.Left + 100, treeView.BoundingRectangle.Top + 60));

        WhileException(window =>
        {
            AutomationElement bbProject = window.FindFirstDescendant(cf.ByAutomationId("bbProject"));
            MouseClick(new Point(bbProject.BoundingRectangle.Left + 50, bbProject.BoundingRectangle.Top + 100));
        });

        WhileException(window =>
        {
            AutomationElement treeView = window.FindFirstDescendant(cf.ByAutomationId("TreeView1"));
            MouseDoubleClick(new Point(treeView.BoundingRectangle.Left + 100, treeView.BoundingRectangle.Top + 40));
        });

        WhileException(window =>
        {
            AutomationElement buttonBar = window.FindFirstDescendant(cf.ByAutomationId("ButtonBar1"));
            MouseClick(new Point(buttonBar.BoundingRectangle.Left + 50, buttonBar.BoundingRectangle.Top + 12));
        });

        WhileException(window =>
        {
            AutomationElement buttonBar = window.FindFirstDescendant(cf.ByAutomationId("NavigationBar1"));
            MouseClick(new Point(buttonBar.BoundingRectangle.Left + 50, buttonBar.BoundingRectangle.Bottom - 30));
        });

        // WhileException(window =>
        // {
        //     Grid flexGrid = window.FindFirstDescendant(cf.ByAutomationId("ProductGrid")).AsGrid();
        //     foreach (GridRow row in flexGrid.Rows)
        //     {
        //         row.Select();
        //         row.DrawHighlight();
        //     }
        // });
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
