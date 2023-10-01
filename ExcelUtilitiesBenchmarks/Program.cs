// See https://aka.ms/new-console-template for more information

using BenchmarkDotNet.Running;
using ExcelUtilitiesBenchmarks;

var summary = BenchmarkRunner.Run<ExcelToObjectsBenchmark>();
Console.WriteLine(summary);
