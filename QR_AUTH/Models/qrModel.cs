namespace QR_AUTH.Models;

public class QrModel
{
    public class Err
    {
        public int ErrCode { get; set; }
    }

    public class Merchant
    {
        public int Id { get; set; }
        public string Title { get; set; }
        public int State { get; set; }
        public string qr_data { get; set; }
        public string QrLink { get; set; }
    }

    public class Root
    {
        public Err Err { get; set; }
        public List<Merchant> Merchants { get; set; }
    }
}