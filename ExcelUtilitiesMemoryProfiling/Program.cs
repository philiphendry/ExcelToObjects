using ExcelUtilities;

new ExcelToObjects().ReadData<BenchmarkData>(Path.Combine(Environment.CurrentDirectory, @"TestFiles\Benchmark.xlsx"));

[Worksheet(Name = "Sheet1", HasHeadings = true)]
public class BenchmarkData
{
    // ReSharper disable UnusedMember.Local
    [Column] public int First { get; set; }
    [Column] public double Second { get; set; }
    [Column] public string Third { get; set; } = string.Empty;
    [Column] public string Fourth { get; set; } = string.Empty;
    [Column] public DateOnly Fifth { get; set; }
    // ReSharper restore UnusedMember.Local
}
