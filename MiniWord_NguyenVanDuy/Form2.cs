using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MiniWord_NguyenVanDuy
{
    public partial class Form2 : Form
    {
        private RichTextBox richTextBox;  
        private int searchStartIndex = 0;
        public Form2(RichTextBox rtb)
        {
            InitializeComponent();
            this.richTextBox = rtb;
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        // Tìm từ 
        private void btnFindNext_Click(object sender, EventArgs e)
        {
            string searchText = txtFind.Text;

            if (!string.IsNullOrEmpty(searchText))
            {
                if (searchStartIndex >= richTextBox.TextLength)
                {
                    searchStartIndex = 0;
                    richTextBox.SelectAll();  
                    richTextBox.SelectionBackColor = richTextBox.BackColor; 
                }

                int index = richTextBox.Find(searchText, searchStartIndex, RichTextBoxFinds.None);

                if (index != -1)
                {
                    richTextBox.Select(index, searchText.Length); 
                    richTextBox.SelectionBackColor = Color.Yellow; 
                    searchStartIndex = index + searchText.Length;  
                }
                else
                {
                    MessageBox.Show("Không tìm thấy kết quả.", "Kết quả tìm kiếm");
                    searchStartIndex = 0; 
                }
            }
            else
            {
                MessageBox.Show("Vui lòng nhập văn bản cần tìm kiếm.", "Lỗi tìm kiếm");
            }
        }

        // Nút Thay thế
        private void btnReplace_Click(object sender, EventArgs e)
        {
            if (richTextBox.SelectedText == txtFind.Text)
            {
                richTextBox.SelectedText = txtReplace.Text;
            }
            btnFindNext_Click(sender, e);  
        }

        // Nút Thay thế tất cả
        private void btnReplaceAll_Click(object sender, EventArgs e)
        {
            string searchText = txtFind.Text;
            string replaceText = txtReplace.Text;

            if (!string.IsNullOrEmpty(searchText))
            {
                int index = richTextBox.Find(searchText, 0, RichTextBoxFinds.None);

                while (index != -1)
                {
                    richTextBox.Select(index, searchText.Length); 
                    richTextBox.SelectedText = replaceText;       
                    index = richTextBox.Find(searchText, index + replaceText.Length, RichTextBoxFinds.None); 
                }

                MessageBox.Show("Đã thay thế tất cả các kết quả.", "Thay thế tất cả");
            }
        }
    }
}
