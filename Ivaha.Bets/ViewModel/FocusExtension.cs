using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace Ivaha.Bets.ViewModel
{
    public static class FocusExtension
    {
        public  static  bool    GetIsFocused    (DependencyObject obj)              =>  (bool)obj.GetValue(IsFocusedProperty);
        public  static  void    SetIsFocused    (DependencyObject obj, bool value)  =>  obj.SetValue(IsFocusedProperty, value);

        public  static  readonly    DependencyProperty  IsFocusedProperty           =   DependencyProperty.RegisterAttached(
                                                                                        "IsFocused", typeof (bool), typeof (FocusExtension), 
                                                                                        new UIPropertyMetadata(false, OnIsFocusedPropertyChanged));

        private static  void    OnIsFocusedPropertyChanged  (DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (!(d is UIElement uie) || !(e.NewValue is bool boolVal))
                return;

            if (boolVal)
                uie.Focus();

            if (uie is TextBox tb)
            {
                if (boolVal)
                    tb.CaretIndex           =   tb.Text.Length;

                tb.IsReadOnlyCaretVisible   =   boolVal;
            }
        }
    }
}
