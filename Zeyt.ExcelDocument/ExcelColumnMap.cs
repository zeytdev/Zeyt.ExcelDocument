using System.Linq.Expressions;
using System.Reflection;

namespace Zeyt.ExcelDocument
{
    public class ExcelColumnMap<TClass> where TClass : class
    {
        public ExcelColumnMapData<TClass> ExcelColumnMapData { get; }

        public ExcelColumnMap(PropertyInfo property)
        {
            ExcelColumnMapData = new ExcelColumnMapData<TClass>(property);
        }

        public ExcelColumnMap<TClass> Name(string name)
        {
            ExcelColumnMapData.Name = name;
            return this;
        }

        public ExcelColumnMap<TClass> Width(uint width)
        {
            ExcelColumnMapData.Width = width;
            return this;
        }

        public ExcelColumnMap<TClass> Default(object value)
        {
            ExcelColumnMapData.Default = value;
            return this;
        }

        public ExcelColumnMap<TClass> WriteUsing(Func<TClass, object> convertExpression)
        {
            Expression<Func<TClass, object>> expression = x => convertExpression(x);
            ExcelColumnMapData.WriteUsing = expression.Compile();
            return this;
        }
    }
}
