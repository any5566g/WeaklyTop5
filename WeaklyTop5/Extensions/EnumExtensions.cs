namespace System
{
    [AttributeUsage(AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    public sealed class StringValueAttribute : Attribute
    {
        public StringValueAttribute(string value)
        {
            Value = value;
        }

        public string Value { get; private set; }
    }

    public static partial class EnumExtensions
    {
        public static string StringValue(this Enum enumValue)
        {
            var type = enumValue.GetType();
            var field = type.GetField(enumValue.ToString());
            var attributes = field.GetCustomAttributes(typeof(StringValueAttribute), false);
            return (attributes.Length > 0 ? ((StringValueAttribute)attributes[0]).Value : null);
        }
    }
}
