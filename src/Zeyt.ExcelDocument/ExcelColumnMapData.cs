using System.Reflection;

namespace Zeyt.ExcelDocument
{
    public class ExcelColumnMapData<TClass> where TClass : class
    {
        public string? Name { get; set; }
        public uint Width { get; set; }
        public object? Default { get; set; }
        public Func<TClass, object>? WriteUsing { get; set; }
        public PropertyInfo? Property { get; set; }

        public ExcelColumnMapData(PropertyInfo property)
        {
            Property = property;
        }
    }
}
