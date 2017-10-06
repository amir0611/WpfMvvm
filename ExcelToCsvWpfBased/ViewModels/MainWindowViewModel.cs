using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using ExcelToCsvWpfBased.Annotations;
using System.Collections.ObjectModel;
using System.Windows.Input;
using Microsoft.Win32;
using System.IO;
using System.Windows;

namespace ExcelToCsvWpfBased
{
    internal enum MessageType
    {
        Confirmation,
        Notification
    }

    internal sealed class MainWindowViewModel : INotifyPropertyChanged,IDisposable
    {
        private ExcelData myExcelData;
        
        private Action<string, MessageType> myNotifyAction;
        
        private ObservableCollection<string> mySheets;
        private string myFileName;
        private bool myIsAllSheet;
        private bool myIsSingleSheet;
        private string mySelectedSheet;

        private ICommand myBrowseFileCommand;
        private ICommand mySelectSingleSheetCommand;
        private ICommand myConvertToCsvCommand;
        private ICommand myCancelCommand;
        
        private const string FILE_NAME_PROPERTY = "FileName";
        private const string IS_ALL_SHEET_PROPERTY = "IsAllSheet";
        private const string IS_SINGLE_SHEET_PROPERTY = "IsSingleSheet";
        private const string SELECTED_SHEET_PROPERTY = "SheetName";
        private const string SHEETS_PROPERTY = "Sheets";
        
        private bool myIsNotDisposed = true;

        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Default constructor,this required because we need to use this class in XAML.
        /// It supports Blendability principle, Expression Blend tool can find my view model for binding,
        /// if i provide this constructor
        /// </summary>
        public MainWindowViewModel()
        {
            myExcelData = new ExcelData();

            myBrowseFileCommand = new MainWindowCommand(BrowseFile);
            mySelectSingleSheetCommand = new MainWindowCommand(PopulateSheetNames, IsFileNameProvided);
            myConvertToCsvCommand = new MainWindowCommand(ConvertToCsv, IsReadyToBeConverted);
            myCancelCommand = new MainWindowCommand(Cancel);
        }

        #region Properties Bound To Controls

        public bool IsAllSheet 
        {
            set
            {
                if (value != myIsAllSheet)
                {
                    myIsAllSheet = value;
                    NotifyPropertyChanged(IS_ALL_SHEET_PROPERTY);
                }
            }
            get { return myIsAllSheet; }
        }

        public bool IsSingleSheet
        {
            set
            {
                if (value != myIsSingleSheet)
                {
                    myIsSingleSheet = value;
                    NotifyPropertyChanged(IS_SINGLE_SHEET_PROPERTY);
                }
            }
            get { return myIsSingleSheet; }
        }

        public string SelectedSheet
        {
            set
            {
                if (value != mySelectedSheet)
                {
                    mySelectedSheet = value;
                    NotifyPropertyChanged(SELECTED_SHEET_PROPERTY);
                }
            }
            get { return mySelectedSheet; }
        }

        public ObservableCollection<string> Sheets
        {
            set
            {
                if (value != mySheets)
                {
                    mySheets = value;
                    NotifyPropertyChanged(SHEETS_PROPERTY);
                }
            }
            get { return mySheets; }
        }

        // TODO These mesaage boxes should not be done here, as it is view
        // there are various ways to get rid of this,
        // 1. Interactivity dll of prism
        // 2. using delegates
        public string FileName
        {
            get { return myFileName; }
            set 
            {
                if (value != myFileName)
                {
                    myFileName = value;
                    try
                    {
                        myExcelData.Open(myFileName);
                    }
                    catch (InvalidOperationException)
                    {
                        MessageBox.Show("Please select a proper excel sheet");
                    }
                    catch (DatabaseConnectionException)
                    {
                        MessageBox.Show("Connection issue");
                    }
                    NotifyPropertyChanged(FILE_NAME_PROPERTY);
                }
            }
        }

        #endregion

        #region Commands

        public ICommand BrowseFileCommand
        {
            get
            {
                return myBrowseFileCommand;
            }
        }

        public ICommand SelectSingleSheetCommand 
        {
            get
            {
                return mySelectSingleSheetCommand;
            }
        }

        public ICommand ConvertToCsvCommand
        {
            get
            {
                return myConvertToCsvCommand;
            }
        }

        public ICommand CancelCommand
        {
            get
            {
                return myCancelCommand;
            }
        }

        #endregion

        #region CommandHelpers
        private void BrowseFile()
        {
            var openFileDialog = new OpenFileDialog
            {
                DefaultExt = ".xlsx",
                Filter = "All Excel Files (*.xlsx, *.xls)|*.xlsx;*.xls|All files (*.*)|*.*"
            };

            openFileDialog.ShowDialog();

            if (string.IsNullOrEmpty(openFileDialog.FileName))
            {
                return;
            }

            FileName = openFileDialog.FileName;
        }

        private void PopulateSheetNames()
        {
            Sheets = new ObservableCollection<string>(myExcelData.GetSheetNames());
        }

        private bool IsFileNameProvided()
        {
            if (string.IsNullOrEmpty(myFileName))
            {
                return false;
            }

            return true;
        }

        private void ConvertToCsv()
        {
            if (IsAllSheet)
            {
                foreach (var sheetName in Sheets)
                {
                    WriteWorksheetToFile(sheetName);
                }
            }
            else if(IsSingleSheet)
            {
                WriteWorksheetToFile(SelectedSheet);
            }

            ////string path = string.IsNullOrEmpty(MyOutputPathTextBox.Text)
            ////    ? Environment.CurrentDirectory
            ////    : MyOutputPathTextBox.Text;
        }

        private bool IsReadyToBeConverted()
        {
            if (string.IsNullOrEmpty(FileName))
            {
                return false;
            }

            if (IsAllSheet)
            {
                return true;
            }

            return (IsSingleSheet && !string.IsNullOrEmpty(SelectedSheet));
        }

        private void Cancel()
        {
 
        }

        #endregion

        [NotifyPropertyChangedInvocator]
        private void NotifyPropertyChanged(string propertyName)
        {
            var handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        private void WriteWorksheetToFile(string sheetName)
        {
            var csvFileName = sheetName + ".csv";
            //if (!string.IsNullOrEmpty(outputTexxtbox))
            //{
            //    csvFileName = Path.Combine(myFileName, csvFileName);
            //}

            var streamWriter = new StreamWriter(csvFileName);
            myExcelData.GetSheetData(streamWriter, sheetName);
            streamWriter.Close();
        }


        public void Dispose()
        {
            if (myIsNotDisposed)
            {
                myExcelData.Dispose();
                mySheets.Clear();

                mySheets = null;
                myExcelData = null;

                myIsNotDisposed = false;
            }
        }
    }
}
