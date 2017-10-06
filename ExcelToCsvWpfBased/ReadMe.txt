1. I could have used IMultiValueConverter for the convert command, but in that case i would have need to pass parameters from converter
   to excecute method of command, then my command would have become specific to convert, tightly coupled.
   But using MVVM i am having all the View controls property mapped in viewModel, so there is no need to poss parametrs to command
   This way my Command is very generic and is used for alomost all commands.

   CommandParameter needed to be used like this
    <Button x:Name="MyConvertButton" Content="Convert" Grid.Row="5" Grid.Column="1" Command="{Binding ConvertToCsvCommand}">
            <Button.CommandParameter>
                <MultiBinding Converter="{StaticResource FileNameAndSheetOptionConverter}">
                    <Binding ElementName="MyFileNameTextBox" Path="Text"/>
                    <Binding ElementName="MyAllSheetsRadioButton" Path="IsChecked"/>
                    <Binding ElementName="MySingleSheetRadioButton" Path="IsChecked"/>
                    <Binding ElementName="MySheetsComboBox" Path="SelectedValue"/>
                </MultiBinding>
            </Button.CommandParameter>
        </Button>

2. ViewModel can be instantiated either in XAML as rsource, or in behind code.
   If instantiated in xaml , and if i need to use it , then we need to cast the dataContext as ViewModel.

3. I am calling Open() method in FileName setter, ideally it should not be done like this.

4. i am using message boxes in viewmodel which is very bad, i need use delegates to pass this responsibility to view.

5. 