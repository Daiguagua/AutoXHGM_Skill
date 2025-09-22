using System.Windows;

namespace Skill_Loop
{
    public enum ConditionType
    {
        Color,
        OCR
    }

    public partial class ConditionTypeDialog : Window
    {
        public ConditionType SelectedType { get; private set; }

        public ConditionTypeDialog()
        {
            InitializeComponent();
        }

        private void ColorCondition_Click(object sender, RoutedEventArgs e)
        {
            SelectedType = ConditionType.Color;
            DialogResult = true;
            Close();
        }

        private void OcrCondition_Click(object sender, RoutedEventArgs e)
        {
            SelectedType = ConditionType.OCR;
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}