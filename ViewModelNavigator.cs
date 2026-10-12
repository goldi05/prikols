using System;
using System.Collections.Generic;

namespace ViewModel
{
    public class ViewModelNavigator
    {
        private readonly IReadOnlyList<INavigableViewModel> _pages;

        public ViewModelNavigator(IReadOnlyList<INavigableViewModel> pages)
        {
            _pages = pages ?? throw new ArgumentNullException(nameof(pages));
        }

        public void NavigateTo(INavigableViewModel target)
        {
            if (target == null)
                throw new ArgumentNullException(nameof(target));

            foreach (var page in _pages)
                page.IsVisible = ReferenceEquals(page, target);
        }
    }
}
