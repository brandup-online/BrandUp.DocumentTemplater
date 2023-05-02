using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Abstraction
{
    /// <summary>
    /// Коммада обрабатывающая контекст для 
    /// </summary>
    public interface IContextCommand
    {
        /// <summary>
        /// Название комманды для использования в .docx файле.
        /// </summary>
        public string Name { get; }
        /// <summary>
        /// Выполнение команды
        /// </summary>
        /// <param name="properties">Параметры команды</param>
        /// <param name="dataContext">Контекст команды</param>
        /// <returns><see cref="HandleResult"/></returns>
        public HandleResult Execute(List<string> properties, object dataContext);
    }
}
