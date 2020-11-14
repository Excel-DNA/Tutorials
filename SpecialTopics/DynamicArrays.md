# Excel-DNA and Dynamic Arrays

The support for 'Dynamic Arrays' is a major new Excel feature available in the Office 365 versions of Excel since 2020.
There are many excellent resources available that explore the power of dynamic arrays - I provide a few links below as introduction to the topic.
In this tutorial I will show how dynamic arrays interact with user-defined functions defined in Excel-DNA add-ins.

## Background

For this tutorial I will assume you are already familiar with the fundamentals of Excel-DNA and the dynamic arrays feature in Excel.
For general background on dynamic arrays, there are many excellent introductions available - I suggest a few below.

#### Dynamic Array links

[Excel dynamic arrays, functions and formulas by Svetlana Cheusheva from AbleBits](https://www.ablebits.com/office-addins-blog/2020/07/08/excel-dynamic-arrays-functions-formulas/) provides a great introduction to dynamic arrays.

Some of Microsoft's notes on Dynamic Arrays:
* [Preview of Dynamic Arrays in Excel](https://techcommunity.microsoft.com/t5/excel-blog/preview-of-dynamic-arrays-in-excel)
* [Dynamic arrays and spilled array behavior](https://support.office.com/en-us/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)
* [Dynamic array formulas in non-dynamic aware Excel](https://support.office.com/en-us/article/dynamic-array-formulas-in-non-dynamic-aware-excel-696e164e-306b-4282-ae9d-aa88f5502fa2)
* [Implicit intersection operator: @](https://support.office.com/en-us/article/implicit-intersection-operator-ce3be07b-0101-4450-a24e-c1c999be2b34)

Bill Jelen (Mr. Excel) goes into the topic in great details in an [e-book about Dynamic Arrays](https://www.mrexcel.com/products/excel-dynamic-arrays-straight-to-the-point-2nd-edition/), and 
 a [Youtube video](https://youtu.be/ViSEZLPmRvw) showing how powerful the dynamic arrays and mathcing new formulas are.

[Formula vs Formula2](https://docs.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/range-formula-vs-formula2)

And countless more write-ups and videos on YouTube.

## Dynamic arrays and Excel-DNA user-defined functions

The overall message is that dynamic arrays Excel and Excel-DNA add-ins work very well together.
Making add-in functions 'array-friendly' provide elegant solutions to problems that required awkward workarounds in older Excel versions.

### Returning arrays

I'll start with a simple function that returns an array result - just some strings with an array size that is determined by the input parameters.

```cs
public static object dnaMakeArray(int numRows, int numCols)
{
    object[,] result = new object [rows, columns];
    for (int i = 0; i < rows; i++)
    {
        for (int j = 0; j < columns; j++)
        {
            result[i,j] = $"{i}|{j}";
        }
    }
    return result;
}
```

Used in a dynamic arrays version of Excel, the result will automatically spill to the right size.
As the inputs change (the number of rows and columns) the resulting spill region automatically resizes.

A few things to note:
* Arrays with 0 size produce an error
* Arrays with a single element don't result in a spill region (no shadow / blue border in the cell)
* If there is no room to spill, we get a `#SPILL!` error
* If there 

### Array inputs

Next we look at a simple function that describes its single input value.

```cs
```

The `AddThem` starter function taking two numbers and adding, would look like this.
```cs
public static double AddThem(double val1, double val2)
{
    return val1 + val2;
}
```

Let's build an array-aware version of the `AddThem` starter function.
Excel-DNA helps simplify the function a bit when we make the input parameters of type double[,] or object[,] - even with single values we'll get a 1x1 array, so the processing can be more uniform.

```cs
public static double[,] AddThem(double[,] val1, double[,] val2)
{
    // if the inputs are not the same size, we return throw na exception, which returns #VALUE back to Excel
    int rows1 = val1.GetLength(0);
    int cols1 = val1.GetLength(1);
    int rows2 = val2.GetLength(0);
    int cols2 = val2.GetLength(1);
    
    if (rows1 <> rows2 || cols1 <> cols2)
        throw new ArgumentException("Incompatible array sizes");
    
    double[,] result = new double[rows1, cols1];
    for (int i = 0; i < rows; i++)
    {
        for (int j = 0; j < columns; j++)
        {
            result[i,j] = val1[i,j] + val2[i,j];
        }
    }
    return result;
}
```

**NOTE:** One danger of using `double` input parameters is that Excel will convert empty cells to 0-values.
To be more careful about the exact input types, change the parameter types to `object[,]` and check the input types during processing.

### Implicit intersection

Let's now look at implicit intersection and the @-operator.

### `ExcelReference` inputs and results

There are some cases where we don't need to know the input values, but can provide processing based on the input array size.
An example would be a function 

### COM Object Model - `Range.Formula2` to avoid '@'-formulas; `HasSpill` and `SpillRange`

## Compatibility with non-dynamic arrays Excel versions

### 'Classic' ArrayResizer

### Testing for whether the running Excel instance supports dynamic arrays

```cs
        static bool? _supportsDynamicArrays;  
        [ExcelFunction(IsHidden=true)]
        public static bool SupportsDynamicArrays()
        {
            if (!_supportsDynamicArrays.HasValue)
            {
                try
                {
                    var result = XlCall.Excel(614, new object[] { 1 }, new object[] { true });
                    _supportsDynamicArrays = true;
                }
                catch
                {
                    _supportsDynamicArrays = false;
                }
            }
            return _supportsDynamicArrays.Value;
        }
```
