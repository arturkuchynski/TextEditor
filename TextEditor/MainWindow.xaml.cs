using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace TextEditor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            fontFamilyBox.ItemsSource = Fonts.SystemFontFamilies.OrderBy(family => family.Source); // Order FontFamilies in alphabetical order
            fontFamilyBox.SelectedItem = new FontFamily("Calibri");
            fontSizeBox.ItemsSource = new List<double>() { 4, 6, 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 };
            fontSizeBox.SelectedItem = (double)12;
            colorBox.ItemsSource = typeof(Colors).GetProperties();
            colorBox.SelectedItem = typeof(Colors).GetProperty("Black");

            //foreach (var color in colors)
            //{
            //    Console.WriteLine("# " + color);
            //}
        }
        
        /// <summary>
        /// Open file event handler
        /// </summary>
        private void openBtn_Click(object sender, RoutedEventArgs e)
        {
 
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();


                if (ofd.ShowDialog() == true)
                {

                    ofd.Filter = "Text Files (*.txt)|*.txt|RichText Files (*.rtf)|*.rtf";

                    TextRange file = new TextRange(textBox.Document.ContentStart, textBox.Document.ContentEnd);

                    using (FileStream fs = new FileStream(ofd.FileName, FileMode.Open))
                    {
                        if (Path.GetExtension(ofd.FileName).ToLower() == ".txt")
                            file.Load(fs, DataFormats.Text);

                        if (Path.GetExtension(ofd.FileName).ToLower() == ".rtf")
                            file.Load(fs, DataFormats.Rtf);

                      
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Save file event handler
        /// </summary>
        private void saveBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog sfd = new SaveFileDialog();
                
                if (sfd.ShowDialog() == true)
                {
                    sfd.Filter = "RichText files (*.rtf)|*.rtf|Text files (*.txt)|*.txt";

                    TextRange file = new TextRange(textBox.Document.ContentStart, textBox.Document.ContentEnd);

                    using (FileStream fs = File.Create(sfd.FileName))
                    {

                        if (Path.GetExtension(sfd.FileName).ToLower() == ".txt")
                            file.Save(fs, DataFormats.Text);

                        else if (Path.GetExtension(sfd.FileName).ToLower() == ".rtf")
                            file.Save(fs, DataFormats.Rtf);

                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Print file event handler
        /// </summary>
        private void printBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                PrintDialog pd = new PrintDialog();

                if ((pd.ShowDialog() == true))
                {
                    // Split text on pages with DocumentPaginator
                    pd.PrintDocument((((IDocumentPaginatorSource)textBox.Document).DocumentPaginator), "Print Document");

                    pd.PrintVisual(textBox as Visual, "Print Visual");

                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Font family selection handler
        /// </summary>
        private void fontFamilyBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (fontFamilyBox.SelectedItem != null && textBox!= null)
            {            
                textBox.Selection.ApplyPropertyValue(FontFamilyProperty, fontFamilyBox.SelectedItem);
         
            }
        }

        /// <summary>
        /// Size selection handler
        /// </summary>
        private void fontSizeBox_SelectionChanged(object sender, RoutedEventArgs e)
        {
            
            if (textBox != null)
            {
                double fontSize = textBox.FontSize;

                if (Double.TryParse(fontSizeBox.Text, out fontSize))
                {
                    textBox.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, fontSize);
                }
            }
        }

        /// <summary>
        /// Color selection handler
        /// </summary>
        private void colorBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                Color selectedColor = (Color)(colorBox.SelectedItem as PropertyInfo).GetValue(null, null);
                textBox.Foreground = new SolidColorBrush(selectedColor);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Selected text properties handler
        /// </summary>
        private void textBox_SelectionChanged(object sender, RoutedEventArgs e)
        {
            try
            {
                object obj = textBox.Selection.GetPropertyValue(TextElement.FontWeightProperty);
                btnBold.IsChecked = (obj != DependencyProperty.UnsetValue) && (obj.Equals(FontWeights.Bold));
                obj = textBox.Selection.GetPropertyValue(TextElement.FontStyleProperty);
                btnItalic.IsChecked = (obj != DependencyProperty.UnsetValue) && (obj.Equals(FontStyles.Italic));
                obj = textBox.Selection.GetPropertyValue(Inline.TextDecorationsProperty);
                btnUnderline.IsChecked = (obj != DependencyProperty.UnsetValue) && (obj.Equals(TextDecorations.Underline));

                fontFamilyBox.SelectedItem = textBox.Selection.GetPropertyValue(TextElement.FontFamilyProperty);
                fontSizeBox.Text = textBox.Selection.GetPropertyValue(TextElement.FontSizeProperty).ToString();
               
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        private void textBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            //TextRange textRange = new TextRange(textBox.Document.ContentStart, textBox.Document.ContentEnd);
            //textRange.ApplyPropertyValue(TextElement.ForegroundProperty, System.Windows.Media.Brushes.Blue);
            //textRange.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
        }

    }
}
