using System;
using ExcelDna.Integration;

public static class Functions
{
    [ExcelFunction(Description = "Hello from Excel-DNA")]
    public static string SayHello(string name) => $"Hello {name}";

    // Dynamic arrays friendly: pass 2D ranges in, return 2D array out (spills in modern Excel)
    [ExcelFunction(Description = "Matrix multiply: returns A x B")]
    public static object MatMul(object[,] a, object[,] b)
    {
        int aRows = a.GetLength(0);
        int aCols = a.GetLength(1);
        int bRows = b.GetLength(0);
        int bCols = b.GetLength(1);

        if (aCols != bRows)
            return $"#DIM! A is {aRows}x{aCols}, B is {bRows}x{bCols}";

        double GetDouble(object v)
        {
            if (v is null) return 0.0;
            if (v is double d) return d;
            if (double.TryParse(v.ToString(), out var dd)) return dd;
            throw new ArgumentException("Non-numeric value encountered.");
        }

        var result = new double[aRows, bCols];
        for (int i = 0; i < aRows; i++)
            for (int k = 0; k < aCols; k++)
            {
                double aik = GetDouble(a[i, k]);
                for (int j = 0; j < bCols; j++)
                    result[i, j] += aik * GetDouble(b[k, j]);
            }

        return result;
    }
}

