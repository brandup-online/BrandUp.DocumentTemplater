namespace BrandUp.DocxGenerator
{
    /// <summary>
    /// Типы комманд.
    /// </summary>
    public enum DocumentGeneratorCommandOutputType
    {
        /// <summary>
        /// Комманда без выполнения действий в документе.
        /// </summary>
        None,
        /// <summary>
        /// Комманда замены вызова комманды в документе результатом выполнения комманды.
        /// </summary>
        Content,
        /// <summary>
        /// Комманда вывода списка в документ.
        /// </summary>
        List
    }
}
