using Allocation_Console_App.Entities;
using NPOI.SS.Formula.Functions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace Allocation_Console_App
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ExcelHandler xlsHandler = new ExcelHandler();
            PropService propService = new PropService();
            AllocationService allocationService = new AllocationService();

            AllocationContext context = new AllocationContext();
            AllocationContext testContext = new AllocationContext();

            bool runProgram = true;
            while (runProgram)
            {
                Console.WriteLine();
                Console.WriteLine();

                // List of possible commands
                Console.WriteLine(" Waiting for input. Possible commands are:                                                                ");
                Console.WriteLine("                                                                                                         ");
                Console.WriteLine("     run [filePath]      -- Runs the complete Programm for a given file.                                 ");
                if (context.AllocationResultsFriday != null && context.AllocationResultsSaturday != null)
                {
                    Console.WriteLine("     evaluate            -- Produces evaluation values for the solution and prints thes to the console.");
                }
                Console.WriteLine("                                                                                                         ");
                Console.WriteLine("     or                                                                                                  ");
                Console.WriteLine("                                                                                                         ");
                Console.WriteLine("     test [filePath]     -- Runs the complete Programm for the given Test-File.                          ");
                Console.WriteLine("     read [filePath]     -- reads data from a given file.                                                ");
                if (context.RawGroupData != null)
                {
                    Console.WriteLine("     groups              -- shows the raw Group data.                                                ");
                    Console.WriteLine("     process             -- process the raw Group data.                                              ");

                }
                if (context.OrderedGroupDataFriday != null && context.OrderedGroupDataSaturday != null)
                {
                    Console.WriteLine("     show                -- shows the ordered Group Lists, if there is any.                          ");
                    Console.WriteLine("     allocate            -- runs the allocation algorithmen.                                         ");
                }
                if (context.AllocationResultsFriday != null && context.AllocationResultsSaturday != null)
                {
                    Console.WriteLine("     result              -- writes the results into a .xls file.                                     ");
                }
                Console.WriteLine();
                Console.Write(" Enter a command (or 'exit' to terminate the program):");
                string command = Console.ReadLine();

                Console.WriteLine();

                if (command.Equals("exit", StringComparison.OrdinalIgnoreCase))
                {
                    // Set the flag to false to terminate the program
                    runProgram = false;
                }
                else
                {
                    if (command.Contains("run"))
                    {
                        // Step 1 - Read input und read file data.
                        string path = string.Empty;
                        var parts = command.Split(" ");
                        if (parts.Length > 1)
                        {
                            path = command.Split(" ")[1];
                        }
                        if (string.IsNullOrEmpty(path))
                        {
                            Console.WriteLine("ERROR - Step 1 failed. No filepath found. Please try again.");
                        }
                        else
                        {
                            context = xlsHandler.ReadExcelFileNPOI(path, context);
                            Console.WriteLine("Succeded - Step 1: Read finished.");
                        }

                        // Step 2 - Process the raw data
                        if (context != null && context.RawGroupData != null)
                        {
                            // Teilaufgabe 1 - Sortierung der Gruppen
                            context.ProcessRawGroupData();
                            Console.WriteLine("Succeded - Step 2: Process finished.");
                        }
                        else
                        {
                            Console.WriteLine("ERROR - Step 2 failed. There is no raw group data to show.");
                        }

                        // Step 3  - Allocate the data
                        if (context.OrderedGroupDataFriday != null && context.OrderedGroupDataSaturday != null)
                        {
                            context.TheaterFriday = propService.CreatePropTables();
                            context.TheaterSaturday = propService.CreatePropTables();
                            context = allocationService.Start(context, true);
                            context = allocationService.Start(context, false);

                            Console.WriteLine("Succeded - Step 3: Allocate finished.");
                        }
                        else
                        {
                            Console.WriteLine("ERROR - Step 3 failed. There is no processed data avaiable.");
                        }

                        string resultFile = string.Empty;
                        // Step 4 - Write Result to the new excel.
                        if (context.AllocationResultsFriday != null && context.AllocationResultsSaturday != null && context.AllocationResultsFriday.Any() && context.AllocationResultsSaturday.Any())
                        {
                            resultFile = xlsHandler.WriteExcelFileNPOI(context);
                            Console.WriteLine($"Succeded - Step 4: Result finished. The created file can be found here: {resultFile}");
                        }
                        else
                        {
                            Console.WriteLine("ERROR - Step 4 failed. There is no results avaiable.");
                        }

                        // step 5 - Write raw data to file in new sheet.
                        if (!string.IsNullOrEmpty(resultFile))
                        {
                            resultFile = xlsHandler.WriteRawDataToFileNPOI(resultFile, context);
                            if (string.IsNullOrEmpty(resultFile))
                            {
                                Console.WriteLine($"ERROR - Step 5 failed. The file could not be created");
                            }
                            else
                            {
                                Console.WriteLine($"Succeded - Step 5: WritingRawData finished. The created file can be found here: {resultFile}");
                            }
                        }

                        // step 6 - Evaluate, create Result list
                        if (context != null)
                        {
                            context.EvaluationResultsFriday = new();
                            context.EvaluationResultsSaturday = new();

                            if (context.TheaterFriday != null)
                            {
                                foreach (var table in context.TheaterFriday)
                                {
                                    context.EvaluationResultsFriday.AddRange(table.EvaluteTable());
                                }
                            }

                            if (context.TheaterSaturday != null)
                            {
                                foreach (var table in context.TheaterSaturday)
                                {
                                    context.EvaluationResultsSaturday.AddRange(table.EvaluteTable());
                                }
                            }

                            Console.WriteLine($"Succeded - Step 6: Evaluating the Solution");
                        }
                        else
                        {
                            Console.WriteLine($"ERROR - Step 6 failed. The results could not be created");
                        }

                        // step 7 - Write result lists to file in new sheet.
                        if (!string.IsNullOrEmpty(resultFile))
                        {
                            if (context != null && context.AllocationResultsFriday != null)
                            {
                                resultFile = xlsHandler.WriteResultsToFileNPOI(resultFile, context);
                                if (string.IsNullOrEmpty(resultFile))
                                {
                                    Console.WriteLine($"ERROR - Step 7 failed. Results could not be written to file");
                                }
                            }

                            if (context != null && context.AllocationResultsSaturday != null)
                            {
                                resultFile = xlsHandler.WriteResultsToFileNPOI(resultFile, context, false);
                            }

                            if (string.IsNullOrEmpty(resultFile))
                            {
                                Console.WriteLine($"ERROR - Step 7 failed. Results could not be written to file");
                            }
                            else
                            {
                                Console.WriteLine($"Succeded - Step 7: WritingResults finished. The created file can be found here: {resultFile}");
                            }
                        }

                        // Finished.
                        Console.WriteLine();
                        Console.WriteLine("Run command finished.");
                    }
                    else if (command.Equals("test"))
                    {
                        // Step 1 - Read input und read file data.
                        string path = Path.Combine(Directory.GetCurrentDirectory(), "eingangs_erfassung_test");
                        path = "C:\\Users\\Robert\\Documents\\GitHub\\bachelorthesis\\Material\\eingangs_erfassung_test.xls";
                        if (string.IsNullOrEmpty(path))
                        {
                            Console.WriteLine("TEST: ERROR - Step 1 failed. No filepath found. Please try again.");
                        }
                        else
                        {
                            testContext = xlsHandler.ReadExcelFileNPOI(path, testContext);
                            Console.WriteLine("TEST: Succeded - Step 1: Read finished.");
                        }

                        // Step 2 - Process the raw data
                        if (testContext != null && testContext.RawGroupData != null)
                        {
                            // Teilaufgabe 1 - Sortierung der Gruppen
                            testContext.ProcessRawGroupData();
                            Console.WriteLine("TEST: Succeded - Step 2: Process finished.");
                        }
                        else
                        {
                            Console.WriteLine("TEST: ERROR - Step 2 failed. There is no raw group data to show.");
                        }

                        // Step 3  - Allocate the data
                        if (testContext.OrderedGroupDataFriday != null)
                        {
                            testContext.TheaterFriday = propService.CreatePropTables(true);
                            testContext = allocationService.Start(testContext, true);

                            Console.WriteLine("TEST: Succeded - Step 3: Allocate finished.");
                        }
                        else
                        {
                            Console.WriteLine("TEST: ERROR - Step 3 failed. There is no processed data avaiable.");
                        }

                        string resultFile = string.Empty;
                        // Step 4 - Write Result to the new excel.
                        if (testContext.AllocationResultsFriday != null && testContext.AllocationResultsFriday.Any())
                        {
                            resultFile = xlsHandler.WriteExcelFileNPOI(testContext, true);
                            if (string.IsNullOrEmpty(resultFile))
                            {
                                Console.WriteLine($"TEST: ERROR - Step 4 failed. The file could not be created");
                            }
                            else
                            {
                                Console.WriteLine($"TEST: Succeded - Step 4: Result finished. The created file can be found here: {resultFile}");
                            }
                        }
                        else
                        {
                            Console.WriteLine("TEST: Step 4 failed. There is no results avaiable.");
                        }

                        // step 5 - Write raw data to file in new sheet.
                        if (!string.IsNullOrEmpty(resultFile))
                        {
                            resultFile = xlsHandler.WriteRawDataToFileNPOI(resultFile, testContext);
                            if (string.IsNullOrEmpty(resultFile))
                            {
                                Console.WriteLine($"TEST: ERROR - Step 5 failed. The file could not be created");
                            }
                            else
                            {
                                Console.WriteLine($"TEST: Succeded - Step 5: WritingRawData finished. The created file can be found here: {resultFile}");
                            }
                        }

                        // step 6 - Evaluate, create Result list
                        if (testContext != null)
                        {
                            testContext.EvaluationResultsFriday = new();
                            testContext.EvaluationResultsSaturday = new();

                            if (testContext.TheaterFriday != null)
                            {
                                foreach (var table in testContext.TheaterFriday)
                                {
                                    testContext.EvaluationResultsFriday.AddRange(table.EvaluteTable());
                                }
                            }

                            if (testContext.TheaterSaturday != null)
                            {
                                foreach (var table in testContext.TheaterSaturday)
                                {
                                    testContext.EvaluationResultsSaturday.AddRange(table.EvaluteTable());
                                }
                            }

                            Console.WriteLine($"TEST: Succeded - Step 6: Evaluating the Solution");
                        }
                        else
                        {
                            Console.WriteLine($"TEST: ERROR - Step 6 failed. The results could not be created");
                        }

                        // step 7 - Write result lists to file in new sheet.
                        if (!string.IsNullOrEmpty(resultFile))
                        {
                            if (testContext != null && testContext.AllocationResultsFriday != null)
                            {
                                resultFile = xlsHandler.WriteResultsToFileNPOI(resultFile, testContext);
                                if (string.IsNullOrEmpty(resultFile))
                                {
                                    Console.WriteLine($"TEST: ERROR - Step 7 failed. Results could not be written to file");
                                }
                            }

                            if (testContext != null && testContext.AllocationResultsSaturday != null)
                            {
                                resultFile = xlsHandler.WriteResultsToFileNPOI(resultFile, testContext, false);
                            }

                            if (string.IsNullOrEmpty(resultFile))
                            {
                                Console.WriteLine($"TEST: ERROR - Step 7 failed. Results could not be written to file");
                            }
                            else
                            {
                                Console.WriteLine($"TEST: Succeded - Step 7: WritingResults finished. The created file can be found here: {resultFile}");
                            }
                        }

                        // Finished.
                        Console.WriteLine();
                        Console.WriteLine("Test command finished.");
                    }
                    else if (command.Equals("evaluate"))
                    {
                        if (context != null && context.TheaterFriday != null && context.TheaterSaturday != null)
                        {
                            int occupiedSeats = 0;
                            int amountOfGuests = 0;
                            int totalSeatCount = 0;
                            context.EvaluationResultsFriday = new();
                            foreach (var table in context.TheaterFriday)
                            {
                                context.EvaluationResultsFriday.AddRange(table.EvaluteTable());
                                occupiedSeats += (table.Seats.Count - table.CountFreeSeats());
                                totalSeatCount += table.Seats.Count;
                            }

                            context.EvaluationResultsFriday = context.EvaluationResultsFriday.OrderBy(x => x.Parent.GroupId).ToList();

                            foreach (var result in context.EvaluationResultsFriday)
                            {
                                amountOfGuests += result.Parent.Size;
                            }

                            Console.WriteLine("Results: ");
                            Console.WriteLine();
                            foreach (var result in context.EvaluationResultsFriday)
                            {
                                Console.WriteLine($"Group: {result.Parent.GroupId}, Table: {result.OwnedSeats[0].Parent.TableNumber}, Group-Size: {result.Parent.Size}, Owned Seat: {result.OwnedSeats.Count}, Main Seat: nr.{result.MainSeat.SeatNumber} - seatscore: {result.MainSeat.Score}, AverageScore: {result.AverageScore}.");
                            }
                            Console.WriteLine();
                            Console.WriteLine($"Occupied Seats: {occupiedSeats}");
                            Console.WriteLine($"Amount of actually placed Guests: {occupiedSeats}");
                            Console.WriteLine($"This numbers should be the same, since there cann´t be more guests then occupied seats");
                            Console.WriteLine($"Amount of TotalSeats: {totalSeatCount}");

                            Console.WriteLine();
                            Console.WriteLine();

                            Console.WriteLine("Allocation Results: ");
                            Console.WriteLine();
                            if (context.AllocationResultsFriday != null)
                            {
                                foreach (var result in context.AllocationResultsFriday)
                                {
                                    Console.WriteLine($"Group: {result.Parent.GroupId}, Table: {result.OwnedSeats[0].Parent.TableNumber}, Group-Size: {result.Parent.Size}, Owned Seat: {result.OwnedSeats.Count}, Main Seat: nr.{result.MainSeat.SeatNumber} - seatscore: {result.MainSeat.Score}, AverageScore: {result.AverageScore}.");
                                }
                            }
                            Console.WriteLine();
                            Console.WriteLine($"Results Count: {context.EvaluationResultsFriday.Count}");
                            Console.WriteLine($"Allocation Results Count: {context.AllocationResultsFriday.Count}");
                            Console.WriteLine();
                            Console.WriteLine();

                            var seatNumbers = context.GetAllocatedSeats().OrderBy(x => x).ToList();

                            for (int i = 1; i < seatNumbers.Count; i++)
                            {
                                if (seatNumbers[1] == seatNumbers[i - 1])
                                {
                                    Console.WriteLine($"SeatNumber: {seatNumbers[i]} allocated multiple times!");
                                }
                            }
                            Console.WriteLine($"Amount of seats: {seatNumbers.Count}");

                            int sum = 0;
                            foreach (var group in context.OrderedGroupDataFriday)
                            {
                                sum += group.Size;
                            }
                            Console.WriteLine($"Needed: {sum}");
                        }
                        else
                        {
                            Console.WriteLine("Evaluate command threw Null-Exception.");
                        }

                        Console.WriteLine("Run command finished.");
                    }
                    else if (command.Contains("read", StringComparison.OrdinalIgnoreCase)) // Read data from a given file
                    {
                        string path = string.Empty;
                        var parts = command.Split(" ");
                        if (parts.Length > 1)
                        {
                            path = command.Split(" ")[1];
                        }
                        if (string.IsNullOrEmpty(path))
                        {
                            Console.WriteLine("No filepath found. Please try again.");
                        }
                        else
                        {
                            context = xlsHandler.ReadExcelFileNPOI(path, context);
                        }
                    }
                    else if (command.Contains("groups"))
                    {
                        if (context != null && context.RawGroupData != null)
                        {
                            Console.WriteLine("Groups: ");

                            foreach (var group in context.RawGroupData)
                            {
                                Console.WriteLine($"ID: {group.GroupId}, Name: {group.Name}, Day: {group.Day}, Size: {group.Size}, Score: {group.Score}.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("There is no raw group data to show.");
                        }
                    }
                    else if (command.Equals("process", StringComparison.OrdinalIgnoreCase))
                    {
                        if (context != null && context.RawGroupData != null)
                        {
                            // Teilaufgabe 1 - Sortierung der Gruppen
                            context.ProcessRawGroupData();
                        }
                        else
                        {
                            Console.WriteLine("There is no raw group data to show.");
                        }
                    }
                    else if (command.Equals("show"))
                    {
                        if (context != null && context.OrderedGroupDataSaturday != null && context.OrderedGroupDataFriday != null)
                        {
                            Console.WriteLine("Friday: ");

                            foreach (var group in context.OrderedGroupDataFriday)
                            {
                                Console.WriteLine($"Day: {group.Day}, Size: {group.Size}, ID: {group.GroupId}, Score: {group.Score}, Name: {group.Name}.");
                            }

                            Console.WriteLine();
                            Console.WriteLine();
                            Console.WriteLine("Saturday: ");

                            foreach (var group in context.OrderedGroupDataSaturday)
                            {
                                Console.WriteLine($"Day: {group.Day}, Size: {group.Size}, ID: {group.GroupId}, Score: {group.Score}, Name: {group.Name}.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("There is no raw group data to show.");
                        }
                    }
                    else if (command.Equals("allocate"))
                    {
                        context.TheaterFriday = propService.CreatePropTables();
                        context.TheaterSaturday = propService.CreatePropTables();

                        context = allocationService.Start(context, true);
                        context = allocationService.Start(context, false);

                        Console.WriteLine("Allocate command finished.");
                    }
                    else if (command.Equals("result"))
                    {
                        xlsHandler.WriteExcelFileNPOI(context);
                        Console.WriteLine("Result command finished.");
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(command))
                        {
                            Console.WriteLine($"No command. Please try one of listed commands");

                        }
                        else
                        {
                            Console.WriteLine($"Unknown command: {command}. Please try another command.");
                        }
                    }
                }
                System.Threading.Thread.Sleep(100);
            }
        }
    }
}