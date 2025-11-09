// SilentExcelObserver.cs
namespace DbcParserLib.Excel.Observers
{
    /// <summary>
    /// Minimal no-op observer used when a parser needs an observer that doesn't collect errors.
    /// </summary>
    internal class SilentExcelObserver : IExcelParseFailureObserver
    {
        public string CurrentSheet { get; set; }
        public int CurrentRow { get; set; }

        public void SheetNotFound(string sheetName) { }
        public void SheetHeaderMissing(string sheetName, string expectedHeader) { }
        public void SheetEmpty(string sheetName) { }
        public void InvalidHexId(string value) { }
        public void InvalidInteger(string fieldName, string value) { }
        public void InvalidFloat(string fieldName, string value) { }
        public void InvalidBoolean(string fieldName, string value) { }
        public void InvalidEnum(string fieldName, string value, string[] validValues) { }
        public void MissingRequiredField(string fieldName) { }
        public void ValueTableFormatError(string valueTableString) { }
        public void ValueTableDuplicate(string tableName) { }
        public void NodeNameInvalid(string nodeName) { }
        public void DuplicatedNode(string nodeName) { }
        public void MessageIdInvalid(string messageId) { }
        public void MessageNameInvalid(string messageName) { }
        public void DuplicatedMessage(string messageId) { }
        public void MessageNotFound(string messageId) { }
        public void SignalNameInvalid(string signalName) { }
        public void DuplicatedSignalInMessage(string messageId, string signalName) { }
        public void SignalFormatError(string fieldName, string value) { }
        public void SignalMessageNotFound(string messageId, string signalName) { }
        public void EnvironmentVariableNameInvalid(string name) { }
        public void DuplicatedEnvironmentVariable(string name) { }
        public void EnvironmentVariableNotFound(string name) { }
        public void PropertyDefinitionInvalid(string propertyName) { }
        public void DuplicatedPropertyDefinition(string propertyName) { }
        public void PropertyNotFound(string propertyName) { }
        public void PropertyValueInvalid(string propertyName, string value) { }
        public void PropertyScopeInvalid(string scope) { }
        public void CommentTypeInvalid(string type) { }
        public void CommentScopeInvalid(string scope) { }
        public void ExtraTransmitterMessageNotFound(string messageId) { }
        public void ExtraTransmitterInvalid(string transmitterName) { }
        public void NodeReferenceNotFound(string nodeName, string context) { }
        public void MessageReferenceNotFound(string messageId, string context) { }
        public void SignalReferenceNotFound(string messageId, string signalName, string context) { }
        public void UnexpectedError(string message) { }
        public void Warning(string message) { }
        public void Clear() { }
    }
}
