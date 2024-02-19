using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;



namespace Conact_Book
{
   
    public partial class Form1 : Form
    {
        private int currentRow = 0;
        private ExcelPackage package = null;
        private List<contact> contacts = new List<contact>();
        private void addTestContacts()
        {
            contact Andre = new contact("John", "Doe", "123 Main St", "555-1234");
        }

        private void LoadExcelData(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    MessageBox.Show("File not found.");
                    return;
                }

                // Обновляем текст label5 с названием выбранного файла
                label5.Text = Path.GetFileName(filePath);

                // Открываем файл Excel для чтения
                FileInfo fileInfo = new FileInfo(filePath);
                package = new ExcelPackage(fileInfo); // Убираем using
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                // Показываем первую строку данных по умолчанию
                currentRow = 1;
                DisplayCurrentRowData(worksheet);
                next_Button.Enabled = rowCount > 1;
           
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while loading Excel data: " + ex.Message);
            }
        }


        private void DisplayCurrentRowData(ExcelWorksheet worksheet)
        {
            // Проверяем, что currentRow находится в допустимом диапазоне строк
            if (currentRow >= 1 && currentRow <= worksheet.Dimension.Rows)
            {
                nameTextBox.Text = worksheet.Cells[currentRow, 1].Value?.ToString();
                surnameTextBox.Text = worksheet.Cells[currentRow, 2].Value?.ToString();
                cellPhoneTextBox.Text = worksheet.Cells[currentRow, 3].Value?.ToString();
                addressTextBox.Text = worksheet.Cells[currentRow, 4].Value?.ToString();
                progressLabel.Text = $"{currentRow}/{worksheet.Dimension.Rows}";
            }
            else
            {
                // Если currentRow находится вне допустимого диапазона строк, выводим сообщение об ошибке
                MessageBox.Show("Current row is out of range.");
            }
        }


        private void highLightButtons()
        {
            List<PictureBox> list = new List<PictureBox>();
            list.Add(add_button);
            list.Add(delete_button);
            list.Add(export_Button);
            list.Add(search_Button);
            list.Add(Load_button);
            list.Add(save_Button);
            list.Add(next_Button);
            list.Add(previous_button);
            foreach (PictureBox item in list)
            {
                HighlightPictureBox(item);
            }
        }
        private void HighlightPictureBox(PictureBox pictureBox)
        {
            // Устанавливаем цвет подсветки
            Color highlightColor = Color.LightGray;

            // Обработчик события MouseEnter
            pictureBox.MouseEnter += (sender, e) =>
            {
                pictureBox.BackColor = highlightColor;
            };

            // Обработчик события MouseLeave
            pictureBox.MouseLeave += (sender, e) =>
            {
                pictureBox.BackColor = Color.Transparent; // Сбрасываем цвет подсветки при уходе мыши
            };
        }

     

        public Form1()
        {
            InitializeComponent();
            highLightButtons();
            /*addLabels();*/
        }

        private void add_button_Click(object sender, EventArgs e)
        {
            if (package == null)
            {
                // Если файл Excel не загружен, добавляем контакт в список contacts
                contact newContact = new contact(nameTextBox.Text, surnameTextBox.Text, cellPhoneTextBox.Text, addressTextBox.Text);
                contacts.Add(newContact);
                currentRow = contacts.Count - 1;
                DisplayCurrentContactData();
                MessageBox.Show("New contact added successfully to the local list.");
            }
            else
            {
                // Если файл Excel загружен, добавляем контакт в файл Excel
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int lastRow = worksheet.Dimension.End.Row;

                worksheet.Cells[lastRow + 1, 1].Value = nameTextBox.Text;
                worksheet.Cells[lastRow + 1, 2].Value = surnameTextBox.Text;
                worksheet.Cells[lastRow + 1, 3].Value = cellPhoneTextBox.Text;
                worksheet.Cells[lastRow + 1, 4].Value = addressTextBox.Text;

                MessageBox.Show("New row added successfully to the Excel file.");
            }
        }

        private void delete_button_Click(object sender, EventArgs e)
        {
    
            if (package == null)
            {
                MessageBox.Show("Please load an Excel file first.");
                return;
            }

            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            int selectedRow = currentRow;

            worksheet.DeleteRow(selectedRow);

            MessageBox.Show("Row deleted successfully.");
        }

        private void edit_button_Click(object sender, EventArgs e)
        {
            MessageBox.Show("To edit a contact, please update the information in the fields with the old data.");
        }


        private void search_Button_Click(object sender, EventArgs e)
        {
            string searchName = Microsoft.VisualBasic.Interaction.InputBox("Enter the name to search:", "Search Contact", "");

            if (!string.IsNullOrWhiteSpace(searchName))
            {
                if (package != null)
                {
                    SearchInExcel(searchName);
                }
                else
                {
                    var foundContacts = contacts.Where(contact => contact.Name.Contains(searchName) || contact.Surname.Contains(searchName)).ToList();


                    if (foundContacts.Any())
                    {
                        string message = "Found contacts:\n\n";
                        foreach (var contact in foundContacts)
                        {
                            message += $"Name: {contact.Name}\nSurname: {contact.Surname}\nCell Phone: {contact.CellPhone}\nAddress: {contact.Address}\n\n";
                        }
                        MessageBox.Show(message, "Search Results");
                    }
                    else
                    {
                        MessageBox.Show($"No contacts found with the name '{searchName}'.", "Search Results");
                    }
                }
               
            }
            else
            {
                MessageBox.Show("Please enter a name to search.", "Search Contact");
            }
        }

        private void SearchInExcel(string searchName)
        {
            if (package != null)
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;
                List<string> foundContacts = new List<string>();

               
                for (int row = 2; row <= rowCount; row++) 
                {
                   
                    for (int col = 1; col <= colCount; col++)
                    {
              
                        string cellValue = worksheet.Cells[row, col].Value?.ToString();

                        
                        if (!string.IsNullOrEmpty(cellValue) && cellValue.IndexOf(searchName, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            string contactInfo = $"Name: {worksheet.Cells[row, 1].Value}\n" +
                                                 $"Surname: {worksheet.Cells[row, 2].Value}\n" +
                                                 $"Cell Phone: {worksheet.Cells[row, 3].Value}\n" +
                                                 $"Address: {worksheet.Cells[row, 4].Value}\n";
                            foundContacts.Add(contactInfo);
                            break; // Если хотя бы одна ячейка в строке соответствует критерию поиска, добавляем строку и переходим к следующей строке
                        }
                    }
                }

                if (foundContacts.Any())
                {
                    string message = "Found contacts:\n\n";
                    foreach (var contact in foundContacts)
                    {
                        message += contact + "\n";
                    }
                    MessageBox.Show(message, "Search Results");
                }
                else
                {
                    MessageBox.Show($"No contacts found with the name '{searchName}'.", "Search Results");
                }
            }
            else
            {
                MessageBox.Show("Please load an Excel file first.");
            }
        }


        private void Load_button_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string selectedFilePath = openFileDialog.FileName;
                LoadExcelData(selectedFilePath);
            }
        }

        private void save_Button_Click(object sender, EventArgs e)
        {
            if (package == null)
            {
                MessageBox.Show("Сначала откройте Excel файл");
            }
            else
            {
                try
                {
                    package.Save();
                    MessageBox.Show("File saved successfully.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while saving the Excel file: " + ex.Message);
                }
            }

        }

        private void CreateExcelFile(string filePath)
        {
            FileInfo newFile = new FileInfo(filePath);
            using (ExcelPackage newPackage = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = newPackage.Workbook.Worksheets.Add("Contacts");

                worksheet.Cells[1, 1].Value = "Name";
                worksheet.Cells[1, 2].Value = "Surname";
                worksheet.Cells[1, 3].Value = "Cell Phone";
                worksheet.Cells[1, 4].Value = "Address";

                int row = 2;
                foreach (var contact in contacts)
                {
                    worksheet.Cells[row, 1].Value = contact.Name;
                    worksheet.Cells[row, 2].Value = contact.Surname;
                    worksheet.Cells[row, 3].Value = contact.CellPhone;
                    worksheet.Cells[row, 4].Value = contact.Address;
                    row++;
                }

                newPackage.Save();
            }
        }





        private void previous_button_Click(object sender, EventArgs e)
        {
            if (package != null)
            {
         
                if (currentRow > 1)
                {
                    currentRow--;
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    DisplayCurrentRowData(worksheet);

              
                    next_Button.Enabled = currentRow < worksheet.Dimension.Rows;

             
                }
                else
                {
                    MessageBox.Show("Beginning of file reached.");
                }
            }
            else 
            {
                if (currentRow > 0)
                {
                    currentRow--;
                    DisplayCurrentContactData();

      
                    next_Button.Enabled = true;
                }
                else
                {
                    MessageBox.Show("Beginning of contacts list reached.");
                }
            }
        }


        private void next_Button_Click(object sender, EventArgs e)
        {

            if (package != null)
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

  
                if (currentRow < worksheet.Dimension.Rows)
                {
                    currentRow++;

                    DisplayCurrentRowData(worksheet);
                }
                else
                {
                    MessageBox.Show("End of file reached.");
                }
            }
            else 
            {
          
                if (contacts.Count > 0 && currentRow < contacts.Count - 1)
                {
                    currentRow++;

              
                    DisplayCurrentContactData();
                }
                else
                {
                    MessageBox.Show("End of contacts list reached.");
                }
            }
        }

        private void DisplayCurrentContactData()
        {
            if (contacts.Count > 0 && currentRow >= 0 && currentRow < contacts.Count)
            {
                nameTextBox.Text = contacts[currentRow].Name;
                surnameTextBox.Text = contacts[currentRow].Surname;
                cellPhoneTextBox.Text = contacts[currentRow].CellPhone;
                addressTextBox.Text = contacts[currentRow].Address;
                progressLabel.Text = $"{currentRow + 1}/{contacts.Count}";
            }
        }
        private void export_Button_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string saveFilePath = saveFileDialog.FileName;
                CreateExcelFile(saveFilePath);
                MessageBox.Show("Contacts saved successfully to the new Excel file.");
            }
        }
    }
}
