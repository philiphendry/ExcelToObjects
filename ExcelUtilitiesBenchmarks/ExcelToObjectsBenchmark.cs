﻿using BenchmarkDotNet.Attributes;
using ExcelUtilities;

namespace ExcelUtilitiesBenchmarks;

[MemoryDiagnoser]
public class ExcelToObjectsBenchmark
{
    private byte[]? _spreadsheetBytes;
    private ExcelToObjects? _excelToObjects;

    [GlobalSetup]
    public void Setup()
    {
        _spreadsheetBytes = File.ReadAllBytes(Path.Combine(Environment.CurrentDirectory, @"TestFiles\Benchmark.xlsx"));
        _excelToObjects = new ExcelToObjects();
    }

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
    
    [Benchmark]
    public void ReadBenchmark()
    {
        var stream = new MemoryStream(_spreadsheetBytes!);
        _excelToObjects!.ReadData<BenchmarkData>(stream);
    }
}