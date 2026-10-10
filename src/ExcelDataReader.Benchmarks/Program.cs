using BenchmarkDotNet.Running;
using ExcelDataReader.Benchmarks;

if (args.Length > 0 && args[0] == "--sst-io")
    SharedStringIo.Run();
else if (args.Length > 0 && args[0] is "--sst-generate" or "--sst-memory")
    SharedStringStorage.RunScenario(args);
else
    BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
