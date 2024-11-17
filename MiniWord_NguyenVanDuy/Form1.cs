using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Word = Microsoft.Office.Interop.Word;

namespace MiniWord_NguyenVanDuy
{
    public partial class Form1 : Form
    {
        private Timer typingTimer;
        private const int typingDelay = 500;
        // Undo
        private Stack<string> undoStack = new Stack<string>();
        // Redo
        private Stack<string> redoStack = new Stack<string>();
        private void InitializeUndoRedo()
        {
            typingTimer = new Timer();
            typingTimer.Interval = typingDelay;
            typingTimer.Tick += (s, e) =>
            {
                typingTimer.Stop();

                if (undoStack.Count == 0 || undoStack.Peek() != richTextBox1.Text)
                {
                    undoStack.Push(richTextBox1.Text);
                    redoStack.Clear();
                }
            };

            richTextBox1.TextChanged += (s, e) =>
            {
                if (!typingTimer.Enabled)
                    typingTimer.Start();
                else
                {
                    typingTimer.Stop();
                    typingTimer.Start();
                }
            };

            undoStack.Push(richTextBox1.Text);
        }

        public Form1()
        {
            InitializeComponent();
            InitializeUndoRedo();
            this.KeyPreview = true;

            // Thanh cuộn
            richTextBox1.ScrollBars = RichTextBoxScrollBars.None;  
            vScrollBar1.Visible = true; 
            vScrollBar1.Scroll += vScrollBar1_Scroll;

            // Font
            foreach (FontFamily font in System.Drawing.FontFamily.Families)
            {
                toolStripComboBoxFont.Items.Add(font.Name);
            }

            if (toolStripComboBoxFont.Items.Contains("Times New Roman"))
            {
                toolStripComboBoxFont.SelectedItem = "Times New Roman";
            }
            else
            {
                toolStripComboBoxFont.Text = richTextBox1.Font.Name;
            }

            richTextBox1.Font = new Font("Times New Roman", richTextBox1.Font.Size);

            //Size
            int[] fontSizes = { 8, 10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 48, 56, 72 };
            foreach (int size in fontSizes)
            {
                toolStripComboBoxSize.Items.Add(size.ToString());
            }

            toolStripComboBoxSize.Text = "13";
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        // Key Down
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            // Undo
            if (e.Control && e.KeyCode == Keys.Z)
            {
                btnUndo_Click(sender, e);
                e.Handled = true;
            }
            // Redo
            else if (e.Control && e.KeyCode == Keys.Y)
            {
                btnRedo_Click(sender, e);
                e.Handled = true;
            }
            // Copy
            else if (e.Control && e.KeyCode == Keys.C)
            {
                btnCopy_Click(sender, e);
                e.Handled = true;
            }
            // Paste
            else if (e.Control && e.KeyCode == Keys.V)
            {
                btnPaste_Click(sender, e);
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.X)
            {
                btnCut_Click(sender, e);
                e.Handled = true;
            }
            else if(e.Control && e.KeyCode == Keys.F)
            {
                OpenFindReplaceForm();
            }

        }

        // Thanh cuộn
        private void UpdateScrollBar()
        {
            int maxLine = richTextBox1.GetLineFromCharIndex(richTextBox1.Text.Length - 1);
            int visibleLines = richTextBox1.Height / richTextBox1.Font.Height;

            if (maxLine > visibleLines)
            {
                vScrollBar1.Maximum = maxLine - visibleLines;
                vScrollBar1.Enabled = true;
            }
            else
            {
                vScrollBar1.Enabled = false;
            }
        }

        private void vScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {
            int maxLine = richTextBox1.GetLineFromCharIndex(richTextBox1.Text.Length - 1);
            if (e.NewValue <= maxLine)
            {
                int firstCharIndex = richTextBox1.GetFirstCharIndexFromLine(e.NewValue);
                richTextBox1.SelectionStart = firstCharIndex;
                richTextBox1.ScrollToCaret();
            }
        }

        // Hàm mở form Tìm kiếm và Thay thế
        private void OpenFindReplaceForm()
        {
            Form2 formFindReplace = new Form2(richTextBox1); 
            formFindReplace.Show();
        }

        /* 1. Hệ thống thực đơn chính, thực đơn ngữ cảnh(có hình, phím nóng) (Tạm ổn) */  
        /* 2. Thanh công cụ, thanh trạng thái(có hình, hình thay đổi trạng thái khi chọn) (Tạm ổn) */
        /* 3. Thực hiện được các chức năng trên thực đơn(và cả trên thanh công cụ) */

        //----------------------------------------- 3.1- New, Open, Save, Save AS, Close, Exit: ---------------------------//

        // New
        private int documentCounter = 1;
        private string currentFilePath = null;
        private bool isDocumentSaved = true;
        private void btnNew_Click(object sender, EventArgs e)
        {
            if (!isDocumentSaved)
            {
                var result = MessageBox.Show("Bạn có muốn lưu tài liệu trước khi tạo mới?", "Xác nhận",
                                             MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    btnSave_Click(sender, e);
                }
                else if (result == DialogResult.Cancel)
                {
                    return;
                }
            }
            richTextBox1.Clear();

            documentCounter++;
            lblTitle.Text = $"Document{documentCounter} - Word -  Saved to this PC";

            currentFilePath = null;
            isDocumentSaved = true;
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            isDocumentSaved = false;
        }

        // Open
        private void btnOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word Documents (*.docx)|*.docx|Rich Text Format (*.rtf)|*.rtf|Text Files (*.txt)|*.txt";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;

                if (filePath.EndsWith(".docx"))
                {
                    OpenDocxFile(filePath);
                }
                else if (filePath.EndsWith(".rtf"))
                {
                    richTextBox1.LoadFile(filePath, RichTextBoxStreamType.RichText);
                }
                else if (filePath.EndsWith(".txt"))
                {
                    richTextBox1.LoadFile(filePath, RichTextBoxStreamType.PlainText);
                }

                UpdateTitleLabel(filePath);
            }
        }
        private void OpenDocxFile(string filePath)
        {
            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;

            try
            {
                doc = wordApp.Documents.Open(filePath);

                string docText = doc.Content.Text;

                richTextBox1.Text = docText;
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                }

                wordApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }
        }

        // Lưu (Save)
        private void Save_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(richTextBox1.Text))
            {
                MessageBox.Show("Không có nội dung để lưu.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Word Document (*.docx)|*.docx";
                saveFileDialog.Title = "Save as Word Document";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;
                    SaveAsWordDocument(filePath);
                    UpdateTitleLabel(filePath);
                }
            }
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(richTextBox1.Text))
            {
                MessageBox.Show("Không có nội dung để lưu.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Word Document (*.docx)|*.docx";
                saveFileDialog.Title = "Save as Word Document";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;
                    SaveAsWordDocument(filePath);
                    UpdateTitleLabel(filePath);
                }
            }
        }

        private void SaveAsWordDocument(string filePath)
        {
            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;

            try
            {
                doc = wordApp.Documents.Add();

                Word.Paragraph para = doc.Content.Paragraphs.Add();
                para.Range.Text = richTextBox1.Text;

                doc.SaveAs2(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi lưu file: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                }

                wordApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }
        }

        // Lưu (Save As)
        private void btnSaveAs_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(richTextBox1.Text))
            {
                MessageBox.Show("Không có nội dung để lưu.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Word Document (*.docx)|*.docx|Rich Text Format (*.rtf)|*.rtf|Text Files (*.txt)|*.txt";
                saveFileDialog.Title = "Lưu Tệp";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    if (filePath.EndsWith(".docx"))
                    {
                        SaveAsWordDocument(filePath);
                    }
                    else if (filePath.EndsWith(".rtf"))
                    {
                        richTextBox1.SaveFile(filePath, RichTextBoxStreamType.RichText);
                    }
                    else if (filePath.EndsWith(".txt"))
                    {
                        richTextBox1.SaveFile(filePath, RichTextBoxStreamType.PlainText);
                    }

                    UpdateTitleLabel(filePath);
                }
            }
        }
        private void UpdateTitleLabel(string filePath)
        {
            int maxTitleLength = 15;

            string fileName = System.IO.Path.GetFileName(filePath);

            if (fileName.Length > maxTitleLength)
            {
                fileName = fileName.Substring(0, maxTitleLength) + "...";
            }

            lblTitle.Text = fileName;
        }

        // Close
        private void btnClose_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(richTextBox1.Text))
            {
                DialogResult result = MessageBox.Show("Bạn có muốn lưu các thay đổi trước khi đóng không?", "Xác nhận", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    btnSaveAs_Click(sender, e);
                }
                else if (result == DialogResult.No)
                {
                    this.Close();
                }
                else
                {
                    return;
                }
            }
            else
            {
                this.Close();
            }
        }

        // Exit
        private void btnExit_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn thoát không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                Application.Exit();
            }
            else
            {
                return;
            }
        }

        //---------------------------- 3.2- Undo, Redo, Cut, Copy, Paste: (menu chính và menu ngữ cảnh)--------------------//

        // Undo
        private void btnUndo_Click(object sender, EventArgs e)
        {
            if (undoStack.Count > 1)
            {
                redoStack.Push(undoStack.Pop());
                richTextBox1.Text = undoStack.Peek();
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
            }
        }

        // Redo
        private void btnRedo_Click(object sender, EventArgs e)
        {
            if (redoStack.Count > 0)
            {
                undoStack.Push(redoStack.Pop());
                richTextBox1.Text = undoStack.Peek();
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
            }
        }

        // Cut
        private void btnCut_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(richTextBox1.SelectedText))
            {
                Clipboard.SetText(richTextBox1.SelectedText);
                richTextBox1.SelectedText = "";
            }
            else
            {
                MessageBox.Show("No text selected to cut.", "Cut", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // Copy
        private void btnCopy_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(richTextBox1.SelectedText))
            {
                Clipboard.SetText(richTextBox1.SelectedText);
            }
            else
            {
                MessageBox.Show("No text selected to copy.", "Copy", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // Paste
        private void btnPaste_Click(object sender, EventArgs e)
        {
            if (Clipboard.ContainsText())
            {
                int selectionStart = richTextBox1.SelectionStart;
                richTextBox1.Text = richTextBox1.Text.Insert(selectionStart, Clipboard.GetText());
                richTextBox1.SelectionStart = selectionStart + Clipboard.GetText().Length;
            }
            else
            {
                MessageBox.Show("No text in the clipboard to paste.", "Paste", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        //----------------------------3.3- Tô màu chữ, màu nền, zoom in, zoom out------------------------------------------//

        // Font Color
        private void btnFontColor_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    richTextBox1.SelectionColor = colorDialog.Color;
                    btnFontColor.BackColor = colorDialog.Color;
                }
            }
        }

        // Background
        private void btnMau_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    richTextBox1.SelectionBackColor = colorDialog.Color;
                    btnMau.BackColor = colorDialog.Color;
                }
            }
        }

        // Zoom in
        private int zoomLevel = 100;
        private void UpdateZoomLabel()
        {
            lblZoomLevel.Text = zoomLevel + "%";
        }
        private void btnZoomIn_Click(object sender, EventArgs e)
        {
            if (zoomLevel < 500)  
            {
                zoomLevel += 10;  
                richTextBox1.Font = new Font(richTextBox1.Font.FontFamily, richTextBox1.Font.Size * 1.1f);  
                UpdateZoomLabel();  
            }
        }

        // Zoom out
        private void btnZoomOut_Click(object sender, EventArgs e)
        {
            if (zoomLevel > 50)  
            {
                zoomLevel -= 10;  
                richTextBox1.Font = new Font(richTextBox1.Font.FontFamily, richTextBox1.Font.Size / 1.1f);  
                UpdateZoomLabel(); 
            }
        }

        //---------------------------3.4- Font chữ, size, style( nghiêng, đậm, gạch chân, và kết hợp với nhau)------------//
        // Font
        private void toolStripComboBoxFont_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (toolStripComboBoxFont.SelectedItem != null)
            {
                string selectedFont = toolStripComboBoxFont.SelectedItem.ToString();
                float fontSize = richTextBox1.SelectionFont?.Size ?? richTextBox1.Font.Size;

                if (richTextBox1.SelectionLength > 0)
                {
                    richTextBox1.SelectionFont = new Font(selectedFont, fontSize);
                }
                else
                {
                    richTextBox1.Font = new Font(selectedFont, fontSize);
                }
            }
        }

        // Size
        private void toolStripComboBoxSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (float.TryParse(toolStripComboBoxSize.SelectedItem.ToString(), out float selectedSize))
            {
                if (richTextBox1.SelectionLength > 0)
                {
                    string selectedFontName = richTextBox1.SelectionFont?.Name ?? richTextBox1.Font.Name;

                    richTextBox1.SelectionFont = new Font(selectedFontName, selectedSize);
                }
                else
                {
                    richTextBox1.Font = new Font(richTextBox1.Font.Name, selectedSize);
                }

                richTextBox1.Focus();
            }
        }

        // Nghiêng
        private void btnItalic_Click(object sender, EventArgs e)
        {
            ToggleFontStyle(FontStyle.Italic);
            if (richTextBox1.SelectionFont != null && richTextBox1.SelectionFont.Style.HasFlag(FontStyle.Bold))
            {
                btnItalic.BackColor = Color.LightGray;
            }
            else
            {
                btnItalic.BackColor = SystemColors.Control;
            }
        }

        // Đậm
        private void btnBold_Click(object sender, EventArgs e)
        {
            ToggleFontStyle(FontStyle.Bold);
            if (richTextBox1.SelectionFont != null && richTextBox1.SelectionFont.Style.HasFlag(FontStyle.Bold))
            {
                btnBold.BackColor = Color.LightGray; 
            }
            else
            {
                btnBold.BackColor = SystemColors.Control; 
            }
        }

        // Gạch chân
        private void btnUnderline_Click(object sender, EventArgs e)
        {
            ToggleFontStyle(FontStyle.Underline);
            if (richTextBox1.SelectionFont != null && richTextBox1.SelectionFont.Style.HasFlag(FontStyle.Bold))
            {
                btnUnderline.BackColor = Color.LightGray;
            }
            else
            {
                btnUnderline.BackColor = SystemColors.Control;
            }
        }

        // Hàm chung để cập nhật định dạng văn bản
        private void ToggleFontStyle(FontStyle style)
        {
            if (richTextBox1.SelectionFont != null)
            {
                FontStyle currentStyle = richTextBox1.SelectionFont.Style;

                if (richTextBox1.SelectionFont.Style.HasFlag(style))
                {
                    currentStyle &= ~style;  
                }
                else
                {
                    currentStyle |= style;   
                }

                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, currentStyle);
            }
        }

        //---------------------------3.5- Tìm kiếm, thay thế-------------------------------------------------------------//
        private void btnTimkiem_ThayThe(object sender, EventArgs e)
        {
            Form2 formFindReplace = new Form2(richTextBox1);  
            formFindReplace.Show();
        }

        private void Search_Click(object sender, EventArgs e)
        {
            Form2 formFindReplace = new Form2(richTextBox1);
            formFindReplace.Show();
        }

        // bug 
        private void button2_Click(object sender, EventArgs e)
        {
            Form2 formFindReplace = new Form2(richTextBox1);
            formFindReplace.Show();
        }

        //---------------------------------------------------------------------------------------------------------------//

        /* 4. Chèn hình ảnh vào văn bản */ 

        private void InsertImage(string imagePath)
        {
            if (System.IO.File.Exists(imagePath))
            {
                Image image = Image.FromFile(imagePath);

                Clipboard.SetDataObject(image, true); 
                richTextBox1.Paste(); 
            }
            else
            {
                MessageBox.Show("Image file not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnInsertImage_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                InsertImage(openFileDialog.FileName);
            }
        }

        //---------------------------------------------------------------------------------------------------------------//

        /* 5. Chèn icons(emoji) từ bảng hiển thị */

        private void btnIcons_Click(object sender, EventArgs e)
        {
            Form3 formIcons = new Form3(richTextBox1);
            formIcons.Show();
        }

        //---------------------------------------------------------------------------------------------------------------//
    }
}
