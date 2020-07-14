using System;
using System.IO;
using System.Text;
using System.Windows.Controls;

namespace BCK
{

    // In order to show your console output with TextBoxOutputter in a TextBox
    // use:
    // Console.SetOut(new TextBoxOutputter(YourTextBox));

    public class TextBoxOutputter : TextWriter
    {
        TextBox textBox = null;
        String strBuffer = String.Empty;
        String strFullContent = String.Empty;

        public String getFullContent()
        {
            return strFullContent;
        }

        public TextBoxOutputter(TextBox output)
        {
            textBox = output;
        }

        public override void Write(char value)
        {
            base.Write(value);
            textBox.Dispatcher.BeginInvoke(new Action(() =>
            {
                if (value.Equals('\n'))
                {
                    String strCurrentContent = strFullContent;

                    String strNewContent =
                        DateTime.Now.ToLongTimeString() + " - " +
                        strBuffer +
                        strCurrentContent;

                    textBox.Text = strNewContent;
                    strFullContent = strNewContent;

                    strBuffer = String.Empty;
                }
                else
                {
                    strBuffer += value.ToString();
                }
            }));
        }

        public override Encoding Encoding
        {
            get { return System.Text.Encoding.UTF8; }
        }
    }
}
