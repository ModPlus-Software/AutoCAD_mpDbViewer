using mpPInterface;

namespace mpDBviewer
{
    public class Interface : IPluginInterface
    {
        public string Name => "mpDBviewer";
        public string AvailCad => "2018";
        public string LName => "Нормативная база";
        public string Description => "Функция для просмотра нормативной базы данных плагина";
        public string Author => "Пекшев Александр aka Modis";
        public string Price => "0";
    }
}
