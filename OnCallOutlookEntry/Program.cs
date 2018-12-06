using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using OfficeOpenXml;

namespace OnCallOutlookEntry
{
    class Program
    {
        static async Task<int> Main(string[] args)
        {
            string filepath;
            string name;
            string email;
            SecureString password;

            if (!args.Any() || args.Length < 4)
            {
                PrintUsage();
                //throw new ArgumentException("See usage", nameof(args));
                Console.Write("Path to xlsm: ");
                filepath = Console.ReadLine();

                Console.Write("Your Name: ");
                name = Console.ReadLine();

                Console.Write("Your email: ");
                email = Console.ReadLine();

                Console.Write("AD password: ");
                password = ReadPassword();
            }
            else
            {
                filepath = args[0];
                name = args[1];
                email = args[2];
                password = args[3].Aggregate(new SecureString(), (ss, c) => { ss.AppendChar(c); return ss; });
                password.MakeReadOnly();
            }

            if (!File.Exists(filepath))
            {
                throw new FileNotFoundException("The file can not be found.", filepath);
            }

            Console.WriteLine();

            string amsName;
            var list = new Dictionary<DateTime, DateTime>();

            DateTime? firstDate = null;
            DateTime? lastDate = null;

            using (var package = new ExcelPackage(new FileInfo(filepath)))
            {
                amsName = package.Workbook.Worksheets["Settings"].GetValue<string>(4, 4);

                var sheet = package.Workbook.Worksheets["OnCall Plan"];

                var range = sheet.Tables["tblPlan"].Address;

                var dateColumn = range.Start.Column + 1;
                var dataStart = range.Start.Column + 2;
                var dataEnd = dataStart + 7;

                for (int row = range.Start.Row + 1; row < range.Start.Row + range.Rows; row++)
                {
                    if (!sheet.Row(row).Hidden)
                    {
                        var weekStart = sheet.GetValue<DateTime>(row, dateColumn);
                        if (!firstDate.HasValue)
                        {
                            firstDate = weekStart;
                        }
                        lastDate = weekStart.AddDays(7);

                        for (int col = dataStart; col < dataEnd; col++)
                        {
                            if (!sheet.Column(col).Hidden)
                            {
                                var oncallerName = sheet.GetValue<string>(row, col);
                                if (oncallerName == name)
                                {
                                    var date = weekStart.AddDays(col - dataStart);
                                    if (!list.Any())
                                    {
                                        list.Add(date, date);
                                    }
                                    else
                                    {
                                        var entry = list.Last();
                                        if (date.Subtract(entry.Value) > TimeSpan.FromDays(1))
                                        {
                                            list.Add(date, date);
                                        }
                                        else
                                        {
                                            list[entry.Key] = date;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            Console.WriteLine($"Found {list.Count} blocks between {firstDate.GetValueOrDefault().ToShortDateString()} and {lastDate.GetValueOrDefault().ToShortDateString()}.");

            string serverAddress = "https://outlook.office365.com/EWS/Exchange.asmx";
            var service = new ExchangeService(ExchangeVersion.Exchange2013_SP1)
            {
                Url = new Uri(serverAddress),
                Credentials = new NetworkCredential(email, password, "tvdit.onmicrosoft.com"),
            };


            // delete old entries
            // Initialize the calendar folder object with only the folder ID. 
            CalendarFolder calendar = CalendarFolder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());

            // Set the start and end time and number of appointments to retrieve.
            CalendarView cView = new CalendarView(firstDate.Value, lastDate.Value);

            // Limit the properties returned to the appointment's subject, start time, and end time.
            cView.PropertySet = new PropertySet(ItemSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End);

            // Retrieve a collection of appointments by using the calendar view.
            var appointments = calendar.FindAppointments(cView).Where(a => a.Subject == $"AMS: {amsName}").ToList();

            Console.WriteLine($"Found {appointments.Count} Appointments on your calendar from {firstDate.Value.ToShortDateString()} to {lastDate.Value.ToShortDateString()} are:");

            foreach (var appointment in appointments)
            {
                appointment.Delete(DeleteMode.MoveToDeletedItems);
                Console.WriteLine($"Deleted old entry from {appointment.Start.ToShortDateString()} to {appointment.End.ToShortDateString()}");
            }

            // save new entries
            foreach (var kvp in list)
            {
                Appointment appointment = new Appointment(service);
                appointment.Subject = $"AMS: {amsName}";
                appointment.IsAllDayEvent = true;
                appointment.Start = kvp.Key;
                appointment.End = kvp.Value.AddDays(1);
                appointment.LegacyFreeBusyStatus = LegacyFreeBusyStatus.Free;
                appointment.IsReminderSet = false;
                appointment.Categories.Add("AMS");
                appointment.Save(SendInvitationsMode.SendToNone);
                Console.WriteLine($"Saved new entry from {appointment.Start.ToShortDateString()} to {appointment.End.ToShortDateString()}");
            }

            Console.WriteLine("Finished");
            Console.ReadKey(true);

            return await System.Threading.Tasks.Task.FromResult(0);
        }

        private static SecureString ReadPassword()
        {
            var pass = new SecureString();
            ConsoleKeyInfo key;
            // Stops Receving Keys Once Enter is Pressed
            while ((key = Console.ReadKey(true)).Key != ConsoleKey.Enter)
            {
                // Backspace Should Not Work
                if (key.Key != ConsoleKey.Backspace)
                {
                    pass.AppendChar(key.KeyChar);
                    Console.Write("*");
                }
                else
                {
                    if (pass.Length > 0)
                    {
                        Console.Write("\b \b");
                        pass.RemoveAt(pass.Length - 1);
                    }
                }
            }
            pass.MakeReadOnly();
            Console.WriteLine();
            return pass;
        }

        static void PrintUsage()
        {
            var exeName = $"{System.AppDomain.CurrentDomain.FriendlyName}.exe";
            Console.WriteLine($"Usage:\n\t{exeName} \"C:\\Path\\to\\OnCallPlan.xlsx\" \"Oncaller name\" \"Oncaller email\" \"Oncaller AD password\"");
            Console.WriteLine();
            Console.WriteLine($"Example:\n\t{exeName} \"C:\\oncall_plan_zharmas.xlsm\" \"Thomas Bandixen\" \"thomas.bandixen@trivadis.com\" \"******\"");
            Console.WriteLine("\t- The name is used to filter the entries");
            Console.WriteLine("\t- The email is used to create new appointments on the exchange server");
            Console.WriteLine("\t- The password is used to authenticate against the exchange server (it wil NOT be logged or saved anywhere)");
            Console.WriteLine();
        }
    }
}
