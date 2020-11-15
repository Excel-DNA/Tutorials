using System;
using System.Threading;
using ExcelDna.Integration;
using static ExcelDna.Integration.XlCall;


public class Functions
{
    public static double dnaAddThem(double val1, double val2)
    {
        return val1 + val2;
    }

    public static object[,] dnaMakeArray(int rows, int cols)
    {
        object[,] result = new object[rows, cols];
        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < cols; j++)
            {
                result[i, j] = $"{i}|{j}";
            }
        }
        return result;
    }

    public static double dnaAddThemArray(double[,] values)
    {
        var rows = values.GetLength(0);
        var cols = values.GetLength(1);
        var sum = 0.0;

        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < cols; j++)
            {
                sum += values[i, j];
            }
        }
        return sum;
    }

    [ExcelFunction(IsMacroType = true)]
    public static string dnaDescribe([ExcelArgument(AllowReference = true)] object arg)
    {
        if (arg is double)
            return "Double: " + (double)arg;
        else if (arg is string)
            return "String: " + (string)arg;
        else if (arg is bool)
            return "Boolean: " + (bool)arg;
        else if (arg is ExcelError)
            return "ExcelError: " + arg.ToString();
        else if (arg is object[,])
            // The object array returned here may contain a mixture of different types,
            // reflecting the different cell contents.
            return string.Format("Array[{0},{1}]", ((object[,])arg).GetLength(0), ((object[,])arg).GetLength(1));
        else if (arg is ExcelMissing)
            return "<<Missing>>"; // Would have been System.Reflection.Missing in previous versions of ExcelDna
        else if (arg is ExcelEmpty)
            return "<<Empty>>"; // Would have been null
        else if (arg is ExcelReference)
            // Calling xlfRefText here requires IsMacroType=true for this function.
            return "Reference: " + Excel(xlfReftext, arg, true); //  + " = " + ((ExcelReference)arg).GetValue().ToString();
        else
            return "!? Unheard Of ?!";
    }

    // To implement an array version, we need to decide how to deal with various size combinations
    public static object[,] dnaConcatenate(string separator, object[,] val1, object[,] val2)
    {
        int rows1 = val1.GetLength(0);
        int cols1 = val1.GetLength(1);
        int rows2 = val2.GetLength(0);
        int cols2 = val2.GetLength(1);

        if (rows1 == rows2 && cols1 == cols2)
        {
            // Same shapes, operate elementwise
            object[,] result = new object[rows1, cols1];
            for (int i = 0; i < rows1; i++)
            {
                for (int j = 0; j < cols1; j++)
                {
                    result[i, j] = $"{val1[i, j]}{separator}{val2[i, j]}";
                }
            }
            return result;
        }

        if (rows1 > 1)
        {
            // Lots of rows in input1, we'll take its first column only, and take the columns of input2
            var rows = rows1;
            var cols = cols2;

            var output = new object[rows, cols];
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    output[i, j] = $"{val1[i, 0]}{separator}{val2[0, j]}";

            return output;
        }
        else
        {

            // Single row in input1, we'll take its columns, and take the rows from input2
            var rows = rows2;
            var cols = cols1;

            var output = new object[rows, cols];
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    output[i, j] = $"{val1[0, j]}{separator}{val2[i, 0]}";

            return output;
        }
    }

    public static object dnaArrayGetHead([ExcelArgument(AllowReference = true)] object input, int numRows)
    {
        if (input is ExcelReference inputRef)
        {
            var rowFirst = inputRef.RowFirst;
            var rowLast = Math.Min(inputRef.RowFirst + numRows, inputRef.RowLast);
            return new ExcelReference(rowFirst, rowLast, inputRef.ColumnFirst, inputRef.ColumnLast, inputRef.SheetId);
        }
        else if (input is object[,] inputArray)
        {
            var rows = inputArray.GetLength(0);
            var cols = inputArray.GetLength(1);

            var resultRows = Math.Min(rows, numRows);
            var result = new object[resultRows, cols];
            for (int i = 0; i < resultRows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    result[i, j] = inputArray[i, j];
                }
            }
            return result;
        }
        else
        {
            // Just a scalar value
            if (numRows >= 1)
            {
                return input;
            }
        }
        // Otherwise we have an error - return #VALUE!
        return ExcelError.ExcelErrorValue;
    }

    public static object dnaMakeArrayAsync(int delayMs, int rows, int cols)
    {
        var funcName = nameof(dnaMakeArrayAsync);
        var args = new object[] { delayMs, rows, cols };

        return ExcelAsyncUtil.Run(funcName, args, () =>
        {
            Thread.Sleep(delayMs);
            object[,] result = new object[rows, cols];
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    result[i, j] = $"{i}|{j}";
                }
            }
            return result;
        });
    }

    // This helper function is converted from https://github.com/Excel-DNA/Samples/blob/master/ArrayMap/Functions.vb
    [ExcelFunction(Description = "Evaluates the two-argument function for every value in the first and second inputs. " + "Takes a single value and any rectangle, or one row and one column, or one column and one row.")]
    public static object dnaArrayMap2([ExcelArgument(Description = "The function to evaluate - either enter the name without any quotes or brackets (for .xll functions), or as a string (for VBA functions)")] object function, [ExcelArgument(Description = "The input value(s) for the first argument (row, column or rectangular range) ")] object input1, [ExcelArgument(Description = "The input value(s) for the second argument (row, column or rectangular range) ")] object input2)
    {
        {
            Func<object, object, object> evaluate;
            if (function is double)
            {
                evaluate = (x, y) => Excel(xlUDF, function, x, y);
            }
            else if (function is string)
            {
                // First try to get the RegisterId, if it's an .xll UDF
                object registerId;
                registerId = Excel(xlfEvaluate, function);
                if (registerId is double)
                {
                    evaluate = (x, y) => Excel(xlUDF, registerId, x, y);
                }
                else
                {
                    // Just call as string, hoping it's a valid VBA function
                    evaluate = (x, y) => Excel(xlUDF, function, x, y);
                }
            }
            else
            {
                return ExcelError.ExcelErrorValue;
            }

            // Check for the case where one of the arguments is not an array, so we evaluate as a 1D function
            if (!(input1 is object[,] inputArr1))
            {
                object evaluate1(object x) => evaluate(input1, x);
                return ArrayEvaluate(evaluate1, input2);
            }
            if (!(input2 is object[,] inputArr2))
            {
                object evaluate1(object x) => evaluate(x, input2);
                return ArrayEvaluate(evaluate1, input1);
            }

            // Otherwise we now have the function to evaluate, and two arrays
            return ArrayEvaluate2(evaluate, inputArr1, inputArr2);
        }
    }

    private static object[,] ArrayEvaluate2(Func<object, object, object> evaluate, object[,] inputArr1, object[,] inputArr2)
    {

        // Now we know both input1 and input2 are arrays
        // We assume they are 1D, else we'll do our best to combine - the exact rules might be decided more carefully
        if (inputArr1.GetLength(0) > 1)
        {
            // Lots of rows in input1, we'll take its first column only, and take the columns of input2
            var rows = inputArr1.GetLength(0);
            var cols = inputArr2.GetLength(1);

            var output = new object[rows, cols];
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    output[i, j] = evaluate(inputArr1[i, 0], inputArr2[0, j]);

            return output;
        }
        else
        {

            // Single row in input1, we'll take its columns, and take the rows from input2
            var rows = inputArr2.GetLength(0);
            var cols = inputArr1.GetLength(1);

            var output = new object[rows, cols];
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    output[i, j] = evaluate(inputArr1[0, j], inputArr2[i, 0]);

            return output;
        }
    }
    

    private static object ArrayEvaluate(Func<object, object> evaluate, object input)
    {
        if (input is object[,] inputArr)
        {
            var rows = inputArr.GetLength(0);
            var cols = inputArr.GetLength(1);
            var output = new object[rows, cols];

            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    output[i, j] = evaluate(inputArr[i, j]);

            return output;
        }
        else
        {
            return evaluate(input);
        }
    }
}
