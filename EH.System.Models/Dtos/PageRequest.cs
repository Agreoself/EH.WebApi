using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{
    public class PageRequest
    {
        public int PageIndex { get; set; }
        public int PageSize { get; set; }
        public string where { get; set; }
        public string order { get; set; }

        public string? defaultWhere { get; set; }

    }
    public class PageRequest<T> : PageRequest where T : class
    {
        public Expression<Func<T, bool>> GetWhere()
        {
            Expression<Func<T, bool>> expression = t => 1 == 1;
            Dictionary<string, string> conditions = new Dictionary<string, string>();
            if (!string.IsNullOrEmpty(where))
            {
                var whereList = this.where.Split(',');
                foreach (var where in whereList)
                {
                    var condition = where.Split('=');
                    var key = condition[0];
                    var value = condition[1];

                    conditions.Add(key, value);
                }
                ParameterExpression parameter = Expression.Parameter(typeof(T), "t");



                // 构建逻辑组合的表达式
                expression = BuildFilterExpression(parameter, conditions);

                return expression;
            }
            else
            {
                return expression;
            }

        }

        public Expression<Func<T, object>> GetOrder()
        {
            string orderKey = this.order;
            // 构建参数表达式
            ParameterExpression parameter = Expression.Parameter(typeof(T), "t");


            // 构建排序表达式
            Expression<Func<T, object>> sortExpression = BuildSortExpression(parameter, orderKey);
            return sortExpression;
        }

        static Expression<Func<T, bool>> BuildFilterExpression(ParameterExpression parameter, Dictionary<string, string> filters)
        {
            Expression body = null;

            foreach (var filter in filters)
            {
                //Expression.Equal(Expression.Property(parameter, filter.Key), Expression.Constant(filter.Value));
                var fieldChar = filter.Key.ToCharArray();
                fieldChar[0] = char.ToUpper(fieldChar[0]);
                var fieldName = new string(fieldChar);
                //获取字段的属性信息
                var property = typeof(T).GetProperty(fieldName);

                if (property != null)
                {
                    MemberExpression member = Expression.Property(parameter, property);

                    object convertValue = null; 

                    if (property.PropertyType.ToString().Contains("System.Nullable"))
                    {
                        var type = Nullable.GetUnderlyingType(property.PropertyType);
                        convertValue = Convert.ChangeType(filter.Value, type);
                    }
                    else
                    {
                        convertValue = Convert.ChangeType(filter.Value, property.PropertyType);
                    }

                    var constant = Expression.Constant(convertValue, property.PropertyType);

                    Expression filterExpression;

                    if(property.PropertyType == typeof(string))
                    {
                        var isNull = Expression.Equal(member, Expression.Constant(null, typeof(string)));

                        MethodInfo containsMethod = typeof(string).GetMethod("IndexOf", new[] { typeof(string), typeof(StringComparison) });
                        var contains = Expression.Call(member, containsMethod,
                            Expression.Constant(filter.Value, typeof(string)),
                            Expression.Constant(StringComparison.OrdinalIgnoreCase));

                        filterExpression = Expression.AndAlso(Expression.Not(isNull), Expression.NotEqual(contains, Expression.Constant(-1)));

                        //filterExpression = Expression.Call(member, typeof(string).GetMethod("Contains", new[] { typeof(string) }), Expression.Constant($"{filter.Value}", typeof(string)));
                    }
                    else if (property.PropertyType == typeof(DateTime))
                    {
                        // 提取日期的年、月和日
                        var dateProperty = Expression.Property(member, "Date");
                        var dateValue = Expression.Property(constant, "Date");

                        // 比较年、月和日
                        var yearComparison = Expression.Equal(Expression.Property(dateProperty, "Year"), Expression.Property(dateValue, "Year"));
                        var monthComparison = Expression.Equal(Expression.Property(dateProperty, "Month"), Expression.Property(dateValue, "Month"));
                        var dayComparison = Expression.Equal(Expression.Property(dateProperty, "Day"), Expression.Property(dateValue, "Day"));

                        // 组合比较条件
                        filterExpression = Expression.AndAlso(Expression.AndAlso(yearComparison, monthComparison), dayComparison);
                    }
                    else if (property.PropertyType == typeof(DateTime?))
                    {
                        var nullableProperty = Expression.Convert(member, typeof(DateTime?));

                        var targetDate = Expression.Property(nullableProperty, "Value");

                        var hasValue = Expression.IsTrue(Expression.Property(nullableProperty, typeof(DateTime?).GetProperty("HasValue")));

                        var dateOnlySource = Expression.Condition(
                            hasValue,
                            Expression.Property(targetDate, "Date"),
                            Expression.Default(typeof(DateTime))
                        );

                        var dateOnlyTarget = Expression.Constant(((DateTime)constant.Value).Date);

                        filterExpression = Expression.Equal(dateOnlySource, dateOnlyTarget);
                    }
                    else
                    {
                        filterExpression = Expression.Equal(member, constant);
                    }


                    if (body == null)
                    {
                        body = filterExpression;
                    }
                    else
                    {
                        body = Expression.AndAlso(body, filterExpression);
                    }
                }


            }

            return Expression.Lambda<Func<T, bool>>(body, parameter);
        }

        static Expression<Func<T, object>> BuildSortExpression(ParameterExpression parameter, string fieldName)
        {
            MemberExpression field = Expression.PropertyOrField(parameter, fieldName);
            UnaryExpression conversion = Expression.Convert(field, typeof(object));
            return Expression.Lambda<Func<T, object>>(conversion, parameter);
        }
    }
}
