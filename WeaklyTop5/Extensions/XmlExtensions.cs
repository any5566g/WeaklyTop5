namespace System.Xml
{
    public static partial class XmlExtensions
    {
        public static XmlElement AppendElement(this XmlNode parent, string namespaceURI, string qualifiedName)
        {
            var elm = parent.OwnerDocument.CreateElement(qualifiedName, namespaceURI);
            parent.AppendChild(elm);
            return elm;
        }

        public static XmlAttribute AppendAttribute(this XmlNode parent, string name, string value)
        {
            var att = parent.OwnerDocument.CreateAttribute(name);
            att.Value = value;
            parent.Attributes.Append(att);
            return att;
        }
    }
}