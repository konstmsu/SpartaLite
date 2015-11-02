using System.Collections.ObjectModel;

namespace Sparta.ViewModels
{
    public class FXOptionPricerViewModel
    {
        public ObservableCollection<FXForwardViewModel> Trades { get; set; }
    }
}
