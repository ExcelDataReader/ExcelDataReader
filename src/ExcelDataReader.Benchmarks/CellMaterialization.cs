using System.Linq.Expressions;
using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class CellMaterialization
{
    private Func<double, object> _materialize = null!;

    [Params("Number", "Boolean", "Date", "Duration")]
    public string Values { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup()
    {
        var assembly = typeof(ExcelReaderFactory).Assembly;
        var workbookType = assembly.GetType("ExcelDataReader.Core.CommonWorkbook")!;
        var decodedType = assembly.GetType("ExcelDataReader.Core.DecodedCellValue")!;
        var kindType = assembly.GetType("ExcelDataReader.Core.CellValueKind")!;
        var styleType = assembly.GetType("ExcelDataReader.Core.ExtendedFormat")!;
        var workbook = Activator.CreateInstance(workbookType)!;
        int format = Values switch { "Date" => 14, "Duration" => 46, _ => 0 };
        var style = Activator.CreateInstance(styleType, [format])!;
        var number = Expression.Parameter(typeof(double), "value");
        Expression value = Values == "Boolean" ? Expression.NotEqual(number, Expression.Constant(0D)) : number;
        var constructor = decodedType.GetConstructor([value.Type]);
        var decoded = constructor != null
            ? Expression.New(constructor, value)
            : Expression.New(
                decodedType.GetConstructor([typeof(object), kindType, typeof(uint)])!,
                Expression.Convert(value, typeof(object)),
                Expression.Constant(Enum.ToObject(kindType, 0), kindType),
                Expression.Constant(0U));
        var create = workbookType.GetMethod("CreateCell")!;
        var cell = Expression.Call(
            Expression.Constant(workbook),
            create,
            Expression.Constant(0),
            decoded,
            Expression.Constant(style),
            Expression.Constant(null, create.GetParameters()[3].ParameterType),
            Expression.Constant(false));
        _materialize = Expression.Lambda<Func<double, object>>(Expression.Property(cell, "Value"), number).Compile();
    }

    [Benchmark(OperationsPerInvoke = 1024)]
    public object ConvertCells()
    {
        object result = null!;
        for (int i = 0; i < 1024; i++)
            result = _materialize(i & 1);
        return result;
    }
}