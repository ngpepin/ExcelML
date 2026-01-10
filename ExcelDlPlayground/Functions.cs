using System;
using ExcelDna.Integration;

public static class Functions
{
    /// <summary>
    /// Simple greeting helper to verify the add-in is responding to Excel function calls.
    /// </summary>
    /// <param name="name">Name to include in the greeting.</param>
    /// <returns>Personalized greeting string.</returns>
    [ExcelFunction(Description = "Hello from Excel-DNA")]
    public static string SayHello(string name) => $"Hello {name}";

    // Dynamic arrays friendly: pass 2D ranges in, return 2D array out (spills in modern Excel)
    /// <summary>
    /// Multiply two matrices provided as Excel ranges and return the product as a spilled array.
    /// </summary>
    /// <param name="a">Left matrix (rows x cols).</param>
    /// <param name="b">Right matrix (rows x cols).</param>
    /// <returns>Product matrix or #DIM! text when dimensions are incompatible.</returns>
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

