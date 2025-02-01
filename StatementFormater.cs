namespace StatementImporter
{
    internal static class StatementFormater
    {
        public static List<StatementLine> FromatStatment(string rawCSV)
        {
            var processedLines = new List<StatementLine>();

            var firstLine = true;
            foreach (var line in rawCSV.Split('\n'))
            {
                // Skip the header row...
                if (firstLine)
                {
                    firstLine = false;
                    continue;
                }
                // ...and any blank lines (mostly the last one!)
                if (string.IsNullOrEmpty(line))
                {
                    continue;
                }

                var statementLine = processLine(line);
                if (statementLine != null)
                {
                    processedLines.Add(statementLine);
                }
            }

            return processedLines;
        }

        static StatementLine? processLine(string rawLine)
        {
            var columns = rawLine.Split(',');


            if (columns.Length != 4)
            {
                Console.WriteLine(
                    $"Wrong number of parts ({columns.Length}) in line {rawLine}");
                return null;
            }

            // There are 4 columns in each line, but we only care about the first 3
            if (DateTime.TryParse(columns[0], out var date) == false)
            {
                Console.WriteLine($"{columns[0]} is not a valid date format");
                return null;
            }
            var description = columns[1];
            if (decimal.TryParse(columns[2], out var amount) == false)
            {
                Console.WriteLine($"Invalid amount value {columns[2]}");
                return null;
            }

            var statemnetLine = new StatementLine
            {
                Date = date,
                Descritpion = description
            };

            // The download from FristDirect only has a single amount column, we 
            // want to split that in to either a credt or debit column
            if (amount > 0)
            {
                statemnetLine.Credit = amount;
            }
            else
            {
                statemnetLine.Debit = amount;
            }

            return statemnetLine;
        }
    }
}
