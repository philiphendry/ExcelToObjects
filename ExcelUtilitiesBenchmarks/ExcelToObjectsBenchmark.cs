using BenchmarkDotNet.Attributes;
using ExcelUtilities;

namespace ExcelUtilitiesBenchmarks;

[MemoryDiagnoser]
public class ExcelToObjectsBenchmark
{
    private byte[]? _spreadsheetBytes;

    [GlobalSetup]
    public void Setup()
    {
        _spreadsheetBytes = File.ReadAllBytes(Path.Combine(Environment.CurrentDirectory, @"TestFiles\Benchmark.xlsx"));
    }

    [Worksheet(Name = "Sheet1", HasHeadings = true)]
    private class BenchmarkData
    {
        // ReSharper disable UnusedMember.Local
        [Column] public int First { get; set; }
        [Column] public double Second { get; set; }
        [Column] public string Third { get; set; }
        [Column] public string Fourth { get; set; }
        [Column] public DateOnly Fifth { get; set; }
        // ReSharper restore UnusedMember.Local
    }
    
    [Benchmark]
    public void ReadBenchmark()
    {
        var stream = new MemoryStream(_spreadsheetBytes!);
        ExcelToObjects.ReadData<BenchmarkData>(stream);
    }
}