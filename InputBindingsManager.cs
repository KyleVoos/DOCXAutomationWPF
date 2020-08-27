using System.Windows;
using System.Windows.Data;
using System.Windows.Input;

namespace DOCXAutomationWPF
{
    public static class InputBindingsManager
    {
        public static readonly DependencyProperty UpdatePropertySourceWHenEnterPressedProperty = DependencyProperty.RegisterAttached(
            "UpdatePropertySourceWhenEnterPressed", typeof(DependencyProperty), typeof(InputBindingsManager), new PropertyMetadata(null, OnUpdatePropertySourceWhenEnterPressedPropertyChanged));

        static InputBindingsManager () { }

        public static void SetUpdatePropertySourceWhenEnterPressed(DependencyObject @object, DependencyProperty value)
        {
            @object.SetValue(UpdatePropertySourceWHenEnterPressedProperty, value);
        }

        public static DependencyProperty GetUpdatePropertySourceWhenEnterPressed(DependencyObject @object)
        {
            return (DependencyProperty)@object.GetValue(UpdatePropertySourceWHenEnterPressedProperty);
        }

        private static void OnUpdatePropertySourceWhenEnterPressedPropertyChanged(DependencyObject @object, DependencyPropertyChangedEventArgs e)
        {
            UIElement element = @object as UIElement;

            if (element == null)
                return;
            if (e.OldValue != null)
                element.PreviewKeyDown -= HandlePreviewKeyDown;
            if (e.NewValue != null)
                element.PreviewKeyDown += new KeyEventHandler(HandlePreviewKeyDown);
        }

        static void HandlePreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                DoUpdateSource(e.Source);
        }

        static void DoUpdateSource(object source)
        {
            DependencyProperty property = GetUpdatePropertySourceWhenEnterPressed(source as DependencyObject);

            if (property == null)
                return;

            UIElement elt = source as UIElement;

            if (elt == null)
                return;

            BindingExpression binding = BindingOperations.GetBindingExpression(elt, property);

            if (binding != null)
                binding.UpdateSource();
        }
    }
}
