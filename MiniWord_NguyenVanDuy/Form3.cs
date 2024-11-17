using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace MiniWord_NguyenVanDuy
{
    public partial class Form3 : Form
    {
        private RichTextBox _richTextBox; 

        public Form3(RichTextBox richTextBox)
        {
            InitializeComponent();
            _richTextBox = richTextBox; 
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            List<string> emojiList = new List<string>
            {
                "😊", "😂", "😍", "😎", "🤣", "😜", "😏", "😇", "🤩", "😋",
                "🥰", "🤗", "😬", "😅", "😢", "😭", "😔", "😤", "😡", "🥺",
                "😷", "😜", "🤑", "🤓", "🤭", "🤤", "🤮", "🤯", "😳", "🥶",
                "🥵", "🥴", "😵", "😕", "😟", "😯", "😦", "😧", "😮", "😲",
                "😋", "😌", "😺", "🙀", "🧡", "💛", "💚", "💙", "💜", "❤️",
                "💖", "💝", "💓", "💗", "💞", "💕", "💘", "💋", "💍", "👑",
                "🏅", "🎉", "🏆", "🏅", "🎁", "🎠", "🎮", "🎲", "🎯", "🎨",
                "🖼️", "🎶", "🎧", "🎷", "🎻", "🎤", "🎬", "📸", "📷", "🎥",
                "🎞️", "📺", "📱", "💻", "🖥️", "🖨️", "🖱️", "💻", "🖥️", "🖨️",
                "⌨️", "🕹️", "🎮", "🎲", "🧩", "🏸", "🏓", "🏏",
            };

            // Thêm emoji vào FlowLayoutPanel khi form được load
            foreach (var emoji in emojiList)
            {
                flowPanelEmojis.Controls.Add(CreateEmojiButton(emoji));
            }
        }

        private Button CreateEmojiButton(string emoji)
        {
            Button btn = new Button();
            btn.Text = emoji;
            btn.Font = new Font("Segoe UI Emoji", 18); 
            btn.Size = new Size(60, 60);
            btn.BackColor = Color.White;  
            btn.FlatStyle = FlatStyle.Flat;  
            btn.ForeColor = Color.Black;  
            btn.Click += (sender, e) => InsertEmoji(emoji); 
            return btn;
        }

        private void InsertEmoji(string emoji)
        {
            _richTextBox.AppendText(emoji);
        }
    }
}
