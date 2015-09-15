using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace John_Tidy
{
    class Program
    {
        protected static int origRow;
        protected static int origCol;
        static void Main(string[] args)
        {

            int i = 1, exhaust = 5, alloy = 10, cat_conv = 30, steer_wheel = 40, light = 3, amount, amount_exhaust = 0, amount_alloy = 0, amount_cat_conv = 0, amount_steer_wheel = 0, amount_light = 0;
            double exhaust_price = 49.99, alloy_price = 4.99, cat_conv_price = 199.99, steer_wheel_price = 19.99, light_price = 34.99;
                
    

            


            while (i > 0)
            {
                bool repeat = true, repeat_2 = true;
                string d, discount;
                double exhaust_total = 0, alloy_total = 0, cat_conv_total = 0, steer_wheel_total = 0, light_total = 0, total;
               int x = 1, y = 1;
                Console.Clear();
                Console.SetCursorPosition(origCol + 15, origRow + 5);
                Console.WriteLine("  WELCOME TO VEHICLE PARTS IMPORTER IRELAND LTD. \n\n\t\t\t    Customer    Staff");

                Console.SetCursorPosition(origCol + 37, origRow + 8);
                

                while (x > 0)
                {
                    var w = Console.ReadKey().Key;
                    if (w == ConsoleKey.LeftArrow)
                    {
                        Console.SetCursorPosition(origCol + 31, origRow + 8);
                    }

                    else if (w == ConsoleKey.RightArrow)
                    {
                        Console.SetCursorPosition(origCol + 42, origRow + 8);

                    }

                    else
                    {
                        x = x - 1;
                    }
                }
                Console.SetCursorPosition(origCol + 37, origRow + 8);
               var p = Console.ReadKey().Key;

                if (p == ConsoleKey.LeftArrow)
                {
                    Excel.Application excelApp = new Excel.Application();
                    string myPath = @"C:\Users\Simon\Documents\College\DBS\Info systems\Invoice2.xlsx";
                    excelApp.Workbooks.Open(myPath);
                    int rowIndex = 16, colIndex = 2, colnumber = colIndex + 1, rowTotal = 23, colTotal = 5, rowCoNum = 7,
                        rowAdd1 = 10, rowAdd2 = 11, rowAdd3 = 12; 
                    while (repeat == true)
                    {



                        Console.Clear();
                        Console.SetCursorPosition(origCol, origRow + 5);
                        Console.WriteLine("Please select from the following products" + 
                             "\n\tExhausts...........................Price: " + "e" + exhaust_price + "   Amount: " + exhaust +
                             "\n\tAlloys.............................Price: " + "e" + alloy_price + "    Amount: " + alloy +
                             "\n\tCatalytic Converters...............Price: " + "e" + cat_conv_price + "  Amount: " + cat_conv +
                             "\n\tSteering Wheels....................Price: " + "e" + steer_wheel_price + "   Amount: " + steer_wheel +
                             "\n\tLights.............................Price: " + "e" + light_price + "   Amount: " + light);
                        
                        d = Console.ReadLine();


                        switch (d)
                        {
                            case "exhausts":
                                Console.WriteLine("Please enter the number of " + d + " you would like");
                                amount = int.Parse(Console.ReadLine());


                                if (amount <= exhaust)
                                {

                                    Console.WriteLine(amount + " " + d + " added to cart");
                                    Console.ReadLine();
                                    exhaust = exhaust - amount;
                                    exhaust_total = exhaust_price * amount;


                                    excelApp.Cells[rowIndex, colIndex] = d;
                                    excelApp.Cells[rowIndex, colnumber] = amount;
                                    excelApp.Cells[rowIndex, colnumber + 1] = exhaust_price;

                                    rowIndex = rowIndex + 1;

                                     amount_exhaust = amount_exhaust + amount;
                                }
                                else if (amount > exhaust)
                                {
                                    
                                    Console.WriteLine("Sorry but your order exceeds current stock.\nPlease email sales@vpiil.ie or call +353-1-3600000 to make a special order");
                                    Console.ReadLine();
                                }
                                break;

                            case "alloys":
                                Console.WriteLine("Please enter the number of " + d + " you would like");
                                amount = int.Parse(Console.ReadLine());
                                if (amount <= alloy)
                                {
                                    Console.WriteLine(amount + " " + d + " added to cart");
                                    Console.ReadLine();
                                    alloy = alloy - amount;
                                    alloy_total = alloy_price * amount;

                                    excelApp.Cells[rowIndex, colIndex] = d;
                                    excelApp.Cells[rowIndex, colnumber] = amount;
                                    excelApp.Cells[rowIndex, colnumber + 1] = alloy_price;

                                    rowIndex = rowIndex + 1;
                                    amount_alloy = amount_alloy + amount;
                                }
                                else if (amount > alloy)
                                {
                                    
                                    Console.WriteLine("Sorry but your order exceeds current stock.\nPlease email sales@vpiil.ie or call +353-1-3600000 to make a special order");
                                    Console.ReadLine();
                                }

                                break;

                            case "cat conv":
                                Console.WriteLine("Please enter the number of " + d + " you would like");
                                amount = int.Parse(Console.ReadLine());
                                if (amount <= cat_conv)
                                {
                                    Console.WriteLine(amount + " " + d + "  added to cart");
                                    Console.ReadLine();
                                    cat_conv = cat_conv - amount;
                                    cat_conv_total = cat_conv_price * amount;

                                    excelApp.Cells[rowIndex, colIndex] = d;
                                    excelApp.Cells[rowIndex, colnumber] = amount;
                                    excelApp.Cells[rowIndex, colnumber + 1] = cat_conv_price;

                                    rowIndex = rowIndex + 1;
                                    amount_cat_conv = amount_cat_conv + amount;
                                }
                                else if (amount > cat_conv)
                                {
                                    
                                    Console.WriteLine("Sorry but your order exceeds current stock.\nPlease email sales@vpiil.ie or call +353-1-3600000 to make a special order");
                                    Console.ReadLine();
                                }

                                break;

                            case "steering wheel":
                                Console.WriteLine("Please enter the number of " + d + " you would like");
                                amount = int.Parse(Console.ReadLine());
                                if (amount <= steer_wheel)
                                {
                                    Console.WriteLine(amount + " " + d + " added to cart");
                                    Console.ReadLine();
                                    steer_wheel = steer_wheel - amount;
                                    steer_wheel_total = steer_wheel_price * amount;

                                    excelApp.Cells[rowIndex, colIndex] = d;
                                    excelApp.Cells[rowIndex, colnumber] = amount;
                                    excelApp.Cells[rowIndex, colnumber + 1] = steer_wheel_price;

                                    rowIndex = rowIndex + 1;
                                    amount_steer_wheel = amount_steer_wheel + amount;
                                }
                                else if (amount > steer_wheel)
                                {
                                    
                                    Console.WriteLine("Sorry but your order exceeds current stock.\nPlease email sales@vpiil.ie or call +353-1-3600000 to make a special order");
                                    Console.ReadLine();
                                }

                                break;

                            case "lights":
                                Console.WriteLine("Please enter the number of " + d + " you would like");
                                amount = int.Parse(Console.ReadLine());
                                if (amount <= light)
                                {
                                    Console.WriteLine(amount + " " + d + " added to cart");
                                    Console.ReadLine();
                                    light = light - amount;
                                    light_total = light_price * amount;
                                    excelApp.Cells[rowIndex, colIndex] = d;
                                    excelApp.Cells[rowIndex, colnumber] = amount;
                                    excelApp.Cells[rowIndex, colnumber + 1] = light_price;

                                    rowIndex = rowIndex + 1;
                                    amount_light = amount_light + amount;
                                }
                                else if (amount >= light)
                                {
                                    Console.WriteLine("Sorry but your order exceeds current stock.\nPlease email sales@vpiil.ie or call +353-1-3600000 to make a special order");
                                    Console.ReadLine();
                                }
                                break;

                            default:

                                Console.WriteLine("Sorry but your order exceeds current stock or is not currently available.\nPlease email sales@vpiil.ie or call +353-1-3600000 to make a special order");

                                break;


                        }

                        Console.WriteLine("Would you like to make another order? Y/N");
                        string go = Console.ReadLine();
                        if (go == "y" || go == "Y")
                        {
                            repeat = true;
                        }
                        else if (go == "n" || go == "N")
                        {


                            total = exhaust_total + alloy_total + steer_wheel_total + cat_conv_total + light_total;
                            if (total == 0)
                            {

                                repeat = false;
                            }

                            else if (total > 0)
                            {
                                Console.WriteLine("The total amounts to: " + "e" + string.Format("{0:0.00}", total) + "\nAre you a registered business? Y/N");
                                string account = Console.ReadLine();




                                if (account == "Y" || account == "y")
                                {

                                    while (repeat_2 == true)
                                    {
                                        Console.WriteLine("Please enter your account name:");
                                        string account_name = Console.ReadLine();


                                        if (account_name == "PVO")
                                        {
                                            Console.Clear();
                                            Console.SetCursorPosition(origCol + 30, origRow + 10);
                                            Console.WriteLine("Welcome " + account_name);
                                            repeat_2 = false;
                                        }
                                        else
                                        {
                                            Console.WriteLine("Sorry this isn't a valid account");
                                            repeat_2 = true;
                                        }
                                    }
                                }
                                else if (account == "N" || account == "n")
                                {
                                    int CoNum = 2343;
                                    Console.Clear();
                                    Console.SetCursorPosition(origCol + 19, origRow + 5);
                                    Console.WriteLine("\t  COMPANY REGISTRATION" +
                                        "\n\tCompany name:" +
                                        "\n\taddress1:" +
                                        "\n\taddress2:" +
                                        "\n\taddress3:" +
                                        "\n\towner:");
                                    Console.SetCursorPosition(origCol + 25, origRow + 6);
                                    string name = Console.ReadLine();
                                    Console.SetCursorPosition(origCol + 25, origRow + 7);
                                    string address1 = Console.ReadLine();
                                    Console.SetCursorPosition(origCol + 25, origRow + 8);
                                    string address2 = Console.ReadLine();
                                    Console.SetCursorPosition(origCol + 25, origRow + 9);
                                    string address3 = Console.ReadLine();
                                    Console.SetCursorPosition(origCol + 25, origRow + 10);
                                    string owner = Console.ReadLine();
                                    Console.Clear();
                                    Console.SetCursorPosition(origCol + 19, origRow + 7);
                                    Console.WriteLine("\t  Registeration complete!");
                                    excelApp.Cells[rowCoNum, colIndex] = CoNum;
                                    excelApp.Cells[rowCoNum + 1, colIndex] = name;
                                    excelApp.Cells[rowAdd1, colIndex] = address1;
                                    excelApp.Cells[rowAdd2, colIndex] = address2;
                                    excelApp.Cells[rowAdd3, colIndex] = address3;
                                }
                                Console.ReadLine();
                                if (total >= 200)
                                {
                                    Console.Clear();
                                    Console.SetCursorPosition(origCol, origRow + 7);
                                    Console.WriteLine("Thank you for spending over 200 euro, a 10 percent discount will now be \nimplemented");
                                    total = total - (total * .1);
                                    Console.WriteLine("The total amount is now: " + "e" + string.Format("{0:0.00}", total));
                                    Console.ReadLine();
                                    discount = "10%";
                                    excelApp.Cells[rowTotal - 1, colTotal] = discount;
                                    excelApp.Cells[rowTotal, colTotal] = total;

                                }

                                else if (total < 200)
                                {
                                    discount = "N/A";
                                    excelApp.Cells[rowTotal - 1, colTotal] = discount;
                                }
                                excelApp.Visible = true;


                                repeat = false;
                            }

                        }


                    }
                }
                else if (p == ConsoleKey.RightArrow)
                {
                    Console.Clear();
                    Console.SetCursorPosition(origCol + 15, origRow + 5);
                    Console.WriteLine("\t\t  VPIL STAFF HOMEPAGE");
                    Console.SetCursorPosition(origCol, origRow + 7);
                    Console.WriteLine("Enter password:");
                    string passphrase = Console.ReadLine();

                    if (passphrase == "john_tidy")
                    {

                        Console.Clear();
                        Console.SetCursorPosition(origCol + 16, origRow + 5);
                        Console.WriteLine("\n\n\t\t\t  Stock    Orders");
                        Console.SetCursorPosition(origCol + 33, origRow + 8);

                        while (y > 0)
                        {
                            var g = Console.ReadKey().Key;
                            if (g == ConsoleKey.LeftArrow)
                            {
                                Console.SetCursorPosition(origCol + 25, origRow + 8);
                            }

                            else if (g == ConsoleKey.RightArrow)
                            {
                                Console.SetCursorPosition(origCol + 40, origRow + 8);

                            }

                            else
                            {
                                y = y - 1;
                            }
                        }
                        Console.SetCursorPosition(origCol + 33, origRow + 8);
                      var  t = Console.ReadKey().Key;

                        if (t == ConsoleKey.LeftArrow)
                        {
                            Console.Clear();
                            Console.SetCursorPosition(origCol, origRow + 5);
                            Console.WriteLine("Stock Inventory" +
                                 "\n\tExhausts............................" + exhaust +
                                 "\n\tAlloys............................." + alloy +
                                 "\n\tCatalytic Converters..............." + cat_conv +
                                 "\n\tSteering Wheels...................." + steer_wheel +
                                 "\n\tLights.............................." + light);
                            Console.ReadLine();

                            Console.WriteLine("Press y to update stock");
                            string stock = Console.ReadLine();
                            if (stock == "Y" || stock == "y")
                            {
                                exhaust = exhaust + 10;
                                alloy = alloy + 100;
                                cat_conv = cat_conv + 5;
                                steer_wheel = steer_wheel + 8;
                                light = light + 20;
                                Console.WriteLine("\tExhausts............................" + exhaust +
                                 "\n\tAlloys............................." + alloy +
                                 "\n\tCatalytic Converters..............." + cat_conv +
                                 "\n\tSteering Wheels...................." + steer_wheel +
                                 "\n\tLights.............................." + light +
                                 "\n Stock updated successfully");
                                Console.ReadLine();
                            }

                        }

                        else if (t == ConsoleKey.RightArrow)
                        {
                            Console.Clear();
                            Console.SetCursorPosition(origCol, origRow + 5);
                            Console.WriteLine("Current order" +
                                 "\n\tExhausts...........................Amount: " + amount_exhaust +
                                 "\n\tAlloys.............................Amount: " + amount_alloy +
                                 "\n\tCatalytic Converters...............Amount: " + amount_cat_conv +
                                 "\n\tSteering Wheels....................Amount: " + amount_steer_wheel +
                                 "\n\tLights.............................Amount: " + amount_light + 
                                 "\nPress enter when order is complete");
                            Console.ReadLine();
                            amount_exhaust = 0;
                            amount_alloy = 0;
                            amount_cat_conv = 0;
                            amount_steer_wheel = 0;
                            amount_light = 0;
                        }
                    }
                    else
                    {
                        Console.WriteLine("Password incorrect");
                        Console.ReadLine();
                    }

                }
               
                
            }
        }

    }
}

