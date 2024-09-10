using System.Linq.Expressions;
using System.Reflection;

namespace Zeyt.ExcelDocument
{
    public class ExcelDocumentMap<TClass> where TClass : class
    {
        public List<ExcelColumnMap<TClass>> ExcelColumnMapList { get; } = new List<ExcelColumnMap<TClass>>();

        public ExcelColumnMap<TClass> Map<TProperty>(Expression<Func<TClass, TProperty>> property)
        {
            if (property.Body is UnaryExpression unaryExp)
            {
                if (unaryExp.Operand is MemberExpression memberExp)
                {
                    var propertyInfo = (PropertyInfo)memberExp.Member;
                    var propertyMap = new ExcelColumnMap<TClass>(propertyInfo);
                    ExcelColumnMapList.Add(propertyMap);
                    return propertyMap;
                }
            }
            else if (property.Body is MemberExpression memberExp)
            {
                var propertyInfo = (PropertyInfo)memberExp.Member;
                var propertyMap = new ExcelColumnMap<TClass>(propertyInfo);
                ExcelColumnMapList.Add(propertyMap);
                return propertyMap;
            }

            throw new ArgumentException($"The expression is not valid. [{property}]");
        }
    }
}
