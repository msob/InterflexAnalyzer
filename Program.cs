using OfficeOpenXml;
using System.Globalization;

// License for EPPlus (required for non-commercial use)
ExcelPackage.License.SetNonCommercialOrganization("My Noncommercial organization");

Console.WriteLine("Interflex Analyzer for Monatsjournal Excel Files");
Console.WriteLine("");
Console.WriteLine("Please enter the full or relative path to the Excel file: ");
var filePath = Console.ReadLine()?.Trim().Trim('"', '\'').Trim();
Console.WriteLine();

if (string.IsNullOrWhiteSpace(filePath))
{
    Console.WriteLine("No file path provided. Exiting program.");
    return;
}
else if (!File.Exists(filePath))
{
    Console.WriteLine($"File not found: {filePath}");
    return;
}

var timeBookings = new List<TimeBooking>();
DateTime? lastDate = null;

try
{
    using (var package = new ExcelPackage(new FileInfo(filePath)))
    {
        var worksheet = package.Workbook.Worksheets[0];
        var rowCount = worksheet.Dimension?.Rows ?? 0;

        for (int row = 1; row <= rowCount; row++)
        {
            var hasValidDate = DateTime.TryParse(Convert.ToString(worksheet.Cells[row, 2].Value), out DateTime dateValue);
            var hasValidTime = TimeSpan.TryParse(Convert.ToString(worksheet.Cells[row, 5].Value)?.Trim(' ', '-', '*'), out TimeSpan startTime);

            if (hasValidDate)
            {
                lastDate = dateValue;
            }

            if ((hasValidDate || hasValidTime) && lastDate.HasValue)
            {
                var startTimeText = Convert.ToString(worksheet.Cells[row, 5].Value)?.Trim(' ', '-', '*') ?? "";
                var endTimeText = Convert.ToString(worksheet.Cells[row, 7].Value)?.Trim(' ', '-', '*') ?? "";
                
                TimeSpan? startTimeValue = null;
                if (TimeSpan.TryParse(startTimeText, out TimeSpan parsedStartTime))
                {
                    startTimeValue = parsedStartTime;
                }
                
                TimeSpan? endTimeValue = null;
                if (TimeSpan.TryParse(endTimeText, out TimeSpan parsedEndTime))
                {
                    endTimeValue = parsedEndTime;
                }
                
                var typeValue = Convert.ToString(worksheet.Cells[row, 10].Value) ?? "";
                var ruleViolation = Convert.ToString(worksheet.Cells[row, 12].Value) ?? "";

                timeBookings.Add(new TimeBooking
                {
                    Date = lastDate.Value,
                    StartTime = startTimeValue,
                    EndTime = endTimeValue,
                    Type = typeValue,
                    RuleViolation = ruleViolation
                });
            }
        }
    }

    Console.WriteLine($"Parsed {timeBookings.Count} time booking entries.");

    if (timeBookings.Count == 0)
    {
        Console.WriteLine("No valid entries found in the file.");
        return;
    }

    // Find the year range
    var firstDate = timeBookings.Min(tb => tb.Date);
    var lastBookingDate = timeBookings.Max(tb => tb.Date);
    var year = firstDate.Year;

    // Create list of all days in the year
    var isLeapYear = DateTime.IsLeapYear(year);
    var daysInYear = isLeapYear ? 366 : 365;
    
    var yearStart = new DateTime(year, 1, 1);
    var yearEnd = new DateTime(year, 12, 31);
    
    // Get all unique dates from the bookings
    var uniqueBookingDates = timeBookings.Select(tb => tb.Date.Date).Distinct().ToList();
    
    // Check if all days of the year are covered
    var allDaysInYear = Enumerable.Range(0, daysInYear)
        .Select(offset => yearStart.AddDays(offset))
        .ToList();
    
    var missingDays = allDaysInYear
        .Where(day => !uniqueBookingDates.Contains(day))
        .ToList();
    
    Console.WriteLine();
    Console.WriteLine("Checking, if all Days of the year are present in the parsed bookings:");
    Console.WriteLine($"Total days in year: {daysInYear}");
    Console.WriteLine($"Days with bookings: {uniqueBookingDates.Count}");
    
    if (missingDays.Count == 0)
    {
        Console.WriteLine("Complete year is included in the list.");
    }
    else
    {
        Console.WriteLine($"Incomplete year. Missing {missingDays.Count} days.");
    }
    
    // Month analysis
    Console.WriteLine();
    Console.WriteLine("Checking, if all Days of each month are present in the parsed bookings:");
    for (int month = 1; month <= 12; month++)
    {
        var daysInMonth = DateTime.DaysInMonth(year, month);
        var monthStart = new DateTime(year, month, 1);
        var daysInThisMonth = Enumerable.Range(0, daysInMonth)
            .Select(offset => monthStart.AddDays(offset))
            .ToList();
        
        var daysWithBookingsInMonth = uniqueBookingDates
            .Where(date => date.Month == month && date.Year == year)
            .ToList();
        
        if (daysWithBookingsInMonth.Count == daysInMonth)
        {
            Console.WriteLine($"{CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month)}: Complete");
        }
        else if (daysWithBookingsInMonth.Count > 0)
        {
            Console.WriteLine($"{CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month)}: Partial ({daysWithBookingsInMonth.Count}/{daysInMonth} days)");
        }
        else
        {
            Console.WriteLine($"{CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month)}: No data");
        }
    }
    
    // Bürobuchung = Leerer Typ-String, aber gültige Uhrzeit für Start- und Endzeit
    var officeBookings = timeBookings.Where(p => p.StartTime != null && p.EndTime != null && string.IsNullOrEmpty(p.Type.Trim()));

    // Mobiles Arbeiten = Typ-String "mobiles arbeiten" und gültige Uhrzeit für Start- und Endzeit
    var homeOfficeBookings = timeBookings.Where(p => p.StartTime != null && p.EndTime != null && p.Type.Trim().Equals("mobiles arbeiten", StringComparison.OrdinalIgnoreCase));

    // "Freizeit-Buchung" (sollte nur an arbeitsfreien Tagen existieren) = Leerer Typ-String und keine gültige Uhrzeit für Start- und Endzeit
    var noTimeBookings = timeBookings.Where(p => p.StartTime == null && p.EndTime == null && string.IsNullOrEmpty(p.Type.Trim()) && string.IsNullOrEmpty(p.RuleViolation.Trim()));

    // Max. eine Bürobuchung am Tag zählen.
    var daysWithOfficeBookings = officeBookings.Select(p => p.Date).Distinct();

    // Max. eine Buchung für mobiles Arbeiten am Tag zählen.
    var daysWithHomeOfficeBookings = homeOfficeBookings.Select(p => p.Date).Distinct();

    // Max. eine Freizeitbuchung am Tag zählen.
    var daysWithoutWork = noTimeBookings.Select(p => p.Date).Distinct();

    // Jede Anwesenheit im Büro rechtfertigt eine Anfahrt.
    var legalOfficeDays = daysWithOfficeBookings;

    // Home Office Pauschale gilt nur an Tagen ohne Büroanwesenheit - Tage mit Dienstreisen können spezialfälle sein.
    var legalHomeOfficeDays = daysWithHomeOfficeBookings.Except(daysWithOfficeBookings);
    
    var numberOfLegalOfficeDays = legalOfficeDays.Count();
    var numberOfLegalHomeOfficeDays = legalHomeOfficeDays.Count();
    var numberOfDaysWithoutWork = daysWithoutWork.Count();
    
    var numberOfCathegorizedDays = numberOfLegalOfficeDays + numberOfLegalHomeOfficeDays + numberOfDaysWithoutWork;

    
    // Type analysis - Add this new section
    Console.WriteLine();
    Console.WriteLine("Type Analysis:");
    
    Console.WriteLine($"Days with only 'mobiles arbeiten' bookings:   {numberOfLegalHomeOfficeDays}");
    Console.WriteLine($"Days with at least one in the office booking: {numberOfLegalOfficeDays}");
    Console.WriteLine($"Days without a booking at all:                {numberOfDaysWithoutWork}");
    Console.WriteLine("-----------------------------------------------------------------------------------");
    Console.WriteLine($"Total days as checksum:                       {numberOfCathegorizedDays}");
    Console.WriteLine($"Total days in year:                           {daysInYear}");
    
    if (numberOfCathegorizedDays == daysInYear)
    {
        Console.WriteLine("The number of categorized days matches the total days in the year.");
    }
    else
    {
        Console.WriteLine($"WARNING: The number of categorized days ({numberOfCathegorizedDays}) does NOT match the total days in the year ({daysInYear})!!!");
    }


    // Ask user if they want to see all bookings
    Console.WriteLine();
    Console.Write("Do you want to see all parsed bookings? y/n (default): ");
    var showBookingsInput = Console.ReadLine()?.Trim().ToLower();
    var showBookings = showBookingsInput == "y" || showBookingsInput == "yes";
    
    // Print all bookings only if user wants to see them
    if (showBookings)
    {
        Console.WriteLine();
        Console.WriteLine("All parsed Bookings:");
        Console.WriteLine("----------------------------------------------------------------------------------------");
        Console.WriteLine(" Zuordnung   | Date       | Start  | End    | Type             | Rule Violation");
        Console.WriteLine("----------------------------------------------------------------------------------------");
        
        DateTime lastPrintedDate = timeBookings.OrderBy(b => b.Date).First().Date;
        foreach (var booking in timeBookings.OrderBy(b => b.Date).ThenBy(b => b.StartTime))
        {
            var date = booking.Date.Date;

            if (date != lastPrintedDate)
            {
                Console.WriteLine("----------------------------------------------------------------------------------------");
                lastPrintedDate = date;
            }
            
            var computedDayType = "";
            if (legalOfficeDays.Contains(date))
            {
                computedDayType += "Büro ";
            }
            if (legalHomeOfficeDays.Contains(date))
            {
                computedDayType += "HO ";
            }
            if (daysWithoutWork.Contains(date))
            {
                computedDayType += "Frei";
            }
            
            string startTimeStr = booking.StartTime.HasValue ? booking.StartTime.Value.ToString(@"hh\:mm") : "--:--";
            string endTimeStr = booking.EndTime.HasValue ? booking.EndTime.Value.ToString(@"hh\:mm") : "--:--";
            
            Console.WriteLine($"{computedDayType,-13} | {booking.Date:yyyy-MM-dd} | {startTimeStr} | {endTimeStr} | {booking.Type,-18} | {booking.RuleViolation}");
        }
    }
}
catch (Exception ex)
{
    Console.WriteLine($"Error processing file: {ex.Message}");
}

Console.WriteLine();
Console.WriteLine("Press return to exit...");
Console.ReadLine();

public class TimeBooking
{
    public DateTime Date { get; set; }
    public TimeSpan? StartTime { get; set; }
    public TimeSpan? EndTime { get; set; }
    public string Type { get; set; } = "";
    public string RuleViolation { get; set; } = "";
}
