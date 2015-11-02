using System;

namespace Sparta.ViewModels
{
    public class FXForwardViewModel
    {
        public DateTime SettlementDate { get; set; }
        public decimal DomesticAmount { get; set; }
        public string DomesticCurrency { get; set; }
        public decimal ForeignAmount { get; set; }
        public string ForeignCurrency { get; set; }
    }
}
