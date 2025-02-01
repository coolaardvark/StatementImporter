namespace StatementImporter
{
    internal class StatementLine
    {
        // Defaults to current, but maybe in future I'll be able to add others?
        public string Account { get; set; } = "Current";
        public DateTime Date { get; set; }
        public string Descritpion { get; set; } = null!;
        public decimal Debit { get; set; } = 0m;
        public decimal Credit { get; set; } = 0m;

        public override string ToString()
        {
            return $"{Account}, {Date.ToString("dd/MM/yy")}, " +
                $"{Descritpion}, {Debit}, {Credit}";
        }
    }
}
