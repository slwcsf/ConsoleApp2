// See https://aka.ms/new-console-template for more information
using System.Text.RegularExpressions;
using Xceed.Document.NET;
using Xceed.Words.NET;

Console.WriteLine("Hello, World!");



// 打开已有的DocX文档
using (var doc = DocX.Load("D:\\1.docx"))
{
    // 查找要替换的文本
    string searchText = "[$1]";
    string replaceText = "NewText";

    // 遍历文档中的段落，查找并替换文本
    foreach (var paragraph in doc.Paragraphs)
    {
        if (paragraph.Text.Contains(searchText))
        {

            StringReplaceTextOptions stringReplaceTextOptions = new StringReplaceTextOptions();
            stringReplaceTextOptions.SearchValue= searchText;
            stringReplaceTextOptions.NewValue = replaceText;
            paragraph.ReplaceText(stringReplaceTextOptions);
        }
    }



    // 查找要替换的文本（包括占位符）
     searchText = "[ImagePlaceholder]";

    // 使用替换文本的方法，可以将占位符文本替换为新的文本或图像
    var paragraphImg = doc.Paragraphs.FirstOrDefault(p => p.Text.Contains(searchText));
    if (paragraphImg != null)
    {
        // 使用替换文本的方法，可以将占位符文本替换为新的文本或图像
        doc.ReplaceText(searchText, string.Empty);
        // 清空段落中的内容
        //paragraph.Clear();

        // 插入新的图片
        var image = doc.AddImage("D:\\1.png");
        var picture = image.CreatePicture();


        // 设置段落的对齐方式为居中
        paragraphImg.Alignment = Alignment.center;
        // 可以设置图片的大小、位置等属性
        picture.Width = 200; // 设置图片宽度
        picture.Height = 150; // 设置图片高度

        // 插入图片到该段落
        paragraphImg.InsertPicture(picture);
    }



    // 获取文档中的表格（假设文档中只有一个表格）
    var table = doc.Tables[1]; // 如果有多个表格，请根据需要选择正确的索引
    // 在表格的末尾插入新行
    var newRow = table.InsertRow(table.Rows[table.Rows.Count - 1]);
    // 向新行的单元格中添加内容
    newRow.Cells[0].Paragraphs[0].Append("New Row, Cell 1");
    newRow.Cells[1].Paragraphs[0].Append("New Row, Cell 2");

    // 保存文档
    doc.Save();

}
