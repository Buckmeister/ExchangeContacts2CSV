using System;
using System.Windows;
using System.Windows.Media.Animation;

namespace ExchangeContacts2CSV
{
    /// <summary>
    /// Interaction logic for EditorWindow.xaml
    /// </summary>
    public partial class EditorWindow : Window
    {
        public EditorWindow()
        {
            InitializeComponent();
        }

        public void Align()
        {
            double newLeft = Owner.Left + Owner.Width;
            double newTop = Owner.Top;

            if (Owner.WindowState == WindowState.Maximized)
            {
                newLeft = SystemParameters.PrimaryScreenWidth - Width - 5;
                newTop = 25;
            }

            if (newLeft + Width > SystemParameters.PrimaryScreenWidth)
            {
                newLeft = SystemParameters.PrimaryScreenWidth - Width - 5;
                newTop += 25;
            }

            if (Left == newLeft && Top == newTop) return;

            if (IsVisible)
            {

                DoubleAnimation dblanimFadeOut = new DoubleAnimation(0, TimeSpan.FromSeconds(.1));
                dblanimFadeOut.Completed += (senderFO, argsFO) =>
                {
                    Hide();

                    Left = newLeft;
                    Top = newTop;

                    DoubleAnimation dblanimFadeIn = new DoubleAnimation(1, TimeSpan.FromSeconds(.1));
                    dblanimFadeIn.Completed += (senderFI, argsFI) => Show();
                    this.BeginAnimation(OpacityProperty, dblanimFadeIn);
                };
                this.BeginAnimation(OpacityProperty, dblanimFadeOut);
            }
            else
            {
                Left = newLeft;
                Top = newTop;
            }
        }
    }
}
