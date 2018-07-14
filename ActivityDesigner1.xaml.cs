using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using System.Activities.Presentation.Model;
using System.Activities;

namespace App_Integration
{
    // Interaction logic for ActivityDesigner1.xaml
    public partial class ActivityDesigner1
    {
        public ActivityDesigner1()
        {
            InitializeComponent();
        }

        public InArgument<String> s { get; set; }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select Word File";
            dialog.Filter = "Word Document|*.doc;*.docx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                String filename = System.IO.Path.GetFullPath(dialog.FileName);
                ModelProperty prop = this.ModelItem.Properties["Path1"];
                prop.SetValue(new InArgument<String>(filename));
                
            }
        }
    }
}
