namespace ItemMasterUpdater.Model
{
    internal class ItemData
    {
        public ItemData() { }

        public ItemData(string itemCode, string length, string width, string height, string weight) 
        {
            ItemCode = itemCode;
            Length = length;
            Width = width;
            Height = height;
            Weight = weight;
        }

        public string ItemCode { get; set; }
        public string Length { get; set; }
        public string Width { get; set; }
        public string Height { get; set; }
        public string Weight { get; set; }
    }
}
