using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Abstraction
{
    /// <summary>
    /// Коммада обрабатывающая контекст для записи в элемент управления
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
        /// <param name="parameters">Параметры команды</param>
        /// <param name="dataContext">Контекст команды</param>
        /// <returns><see cref="HandleResult"/></returns>
        public HandleResult Execute(List<string> parameters, object dataContext);
    }
}
