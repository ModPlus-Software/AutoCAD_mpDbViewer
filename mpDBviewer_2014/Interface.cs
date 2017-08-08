using mpPInterface;

namespace mpDBviewer
{
    public class Interface : IPluginInterface
    {
        private const string _Name = "mpDBviewer";
        private const string _AvailCad = "2014";
        private const string _LName = "Нормативная база";
        private const string _Description = "Функция для просмотра нормативной базы данных плагина";
        private const string _Author = "Пекшев Александр aka Modis";
        private const string _Price = "0";
        public string Name { get { return _Name; } }
        public string AvailCad { get { return _AvailCad; } }
        public string LName { get { return _LName; } }
        public string Description { get { return _Description; } }
        public string Author { get { return _Author; } }
        public string Price { get { return _Price; } }

    }
}
