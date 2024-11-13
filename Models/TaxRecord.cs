namespace Assignment.Models
{
    public class TaxRecord
    {
        public int InvNo { get; set; }
        public string InvCURNo { get; set; }
        public DateTime InvDate { get; set; }
        public int CustomerCode { get; set; }
        public string CustomerName { get; set; }
        public string RegCountry { get; set; }
        public decimal TotalValueAfterTaxing { get; set; }
        public decimal TaxingValue { get; set; }
        public decimal TotalValueBeforeTaxing { get; set; }
    }

}

