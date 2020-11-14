# Excel-DNA and Dynamic Arrays

The support for 'Dynamic Arrays' is a major new Excel feature available in the Office 365 versions of Excel since 2020.
There are many excellent resources available that explore the power of dynamic arrays - I provide a few links below as introduction to the topic.
In this tutorial I will show how dynamic arrays interact with user-defined functions defined in Excel-DNA add-ins.

## Background



### Some links

[Excel dynamic arrays, functions and formulas by Svetlana Cheusheva from AbleBits](https://www.ablebits.com/office-addins-blog/2020/07/08/excel-dynamic-arrays-functions-formulas/) provides a great introduction to dynamic arrays.

Microsoft's initial announcements and some discussions on decisions along the way:
* [Preview of Dynamic Arrays in Excel](https://techcommunity.microsoft.com/t5/excel-blog/preview-of-dynamic-arrays-in-excel)
* [Dynamic arrays and spilled array behavior](https://support.office.com/en-us/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)
* [Dynamic array formulas in non-dynamic aware Excel](https://support.office.com/en-us/article/dynamic-array-formulas-in-non-dynamic-aware-excel-696e164e-306b-4282-ae9d-aa88f5502fa2)
* [Implicit intersection operator: @](https://support.office.com/en-us/article/implicit-intersection-operator-ce3be07b-0101-4450-a24e-c1c999be2b34)

Bill Jelen (Mr. Excel) goes into the topic in great details in an [e-book that explains Dyncami Array in detail](https://www.mrexcel.com/products/excel-dynamic-arrays-straight-to-the-point-2nd-edition/), and 
 a [Youtube video](https://youtu.be/ViSEZLPmRvw) showing how powerful the dynamic arrays and mathcing new formulas are.

[Formula vs Formula2](https://docs.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/range-formula-vs-formula2)

And countless more write-ups and videos on YouTube.

## Dynamic arrays and Excel-DNA user-defined functions

### Returning arrays

```cs

```

### Array inputs

```cs
```

#### `#SPILL!` errors

### Implicit intersection

### `ExcelReference` inputs and results

### COM Object Model - `Formula` vs `Formula2`; `HasSpill` and `SpillRange`

## Compatibility with non-dynamic arrays Excel versions

### 'Classic' ArrayResizer


### Testing for whether the running Excel instance supports dynamic arrays

```csharp
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
