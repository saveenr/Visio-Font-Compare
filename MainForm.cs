using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;
using Application = Microsoft.Office.Interop.Visio.Application;

namespace VisioFontCompare
{
    public partial class MainForm : Form
    {
        private static Application _app;
        private List<string> allfontnames;


        string text1 =
    @"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvqxyz1234567890!@#$%^&*()``_-+=[]{}\|;:'"",.<>/?";
        string text2 =
            @"When from behind that craggy steep till then The horizon’s bound, a huge peak, black and huge, As if with voluntary power instinct, Upreared its head. I struck and struck again, And growing still in stature the grim shape Towered up between me and the stars, and still, for so it seemed with purpose of its own And measured motion like a living thing, Strode after me.

-William Wordsworth 
The Prelude, lines 381-389";

        string text3 =
            @"int function_x( int[] a)
{
    string s = ""hello"";
    char c = s[3];
    var o = new SomeObject();
    return 27 * 2;
}
";

        string text4 =
            @"<html>
    <head>
        <title>Untitled</title>
    <head>
    <body>
        <p>Hello World</p>
    </body>
</html>
";

        string stencilname = "basic_u.vss";
        string rectmaster = "Rectangle";

        private VA.Text.CharStyle font1charstyle;
        private VA.Text.CharStyle font2charstyle;

        public MainForm()
        {
            this.InitializeComponent();
            MainForm._app = new Application();
            var docs = MainForm._app.Documents;
            var doc = docs.Add("");
            this.allfontnames = doc.Fonts.AsEnumerable().Select(f => f.Name).OrderBy( s=>s).ToList();

            this.comboBoxFont1.DataSource = this.allfontnames.ToList();
            this.comboBoxFont2.DataSource = this.allfontnames.ToList();

            this.comboBoxFont1.Text = "Arial";
            this.comboBoxFont2.Text = "Calibri";


        }

        void FontCompare()
        {
            this.font1charstyle = VA.Text.CharStyle.None;
            if (this.checkBoxFont1Bold.Checked)
            {
                this.font1charstyle = this.font1charstyle | VA.Text.CharStyle.Bold;
            }
            if (this.checkBoxFont1Italic.Checked)
            {
                this.font1charstyle = this.font1charstyle | VA.Text.CharStyle.Italic;
            }

            this.font2charstyle = VA.Text.CharStyle.None;
            if (this.checkBoxFont2Bold.Checked)
            {
                this.font2charstyle = this.font2charstyle | VA.Text.CharStyle.Bold;
            }
            if (this.checkBoxFont2Italic.Checked)
            {
                this.font2charstyle = this.font2charstyle | VA.Text.CharStyle.Italic;
            }



            var labelfont = "Segoe UI Light";

            var fci = new FontCompareInput();

            var docs = MainForm._app.Documents;
            fci.Document = docs.Add("");
            var dfonts = fci.Document.Fonts;
            fci.TargetFontNames = new List<string> { this.comboBoxFont1.Text, this.comboBoxFont2.Text };
            fci.Styles = new List<VA.Text.CharStyle>();
            fci.Styles.Add(this.font1charstyle);
            fci.Styles.Add(this.font2charstyle);


            fci.FontSizeFormulas = new List<string> { "8pt", "10pt", "12pt", "14pt", "18pt", "28pt" };
            fci.TextBlocks = new[] {this.text1, this.text2, this.text3, this.text4 };
            fci.LabelFontID = dfonts[labelfont].ID;
            fci.TargetFontIDs = fci.TargetFontNames.Select(f => dfonts[f].ID).ToList();

            fci.TargetFontDisplayNames = new List<string>();

            for (int i = 0; i < fci.TargetFontNames.Count; i++)
            {
                if (fci.Styles[i]!=VA.Text.CharStyle.None)
                {
                    var sb = new StringBuilder();
                    sb.Append(fci.TargetFontNames[i]);
                    sb.Append(" (");
                    var tokens = new List<string>();
                    if ((fci.Styles[i] & VA.Text.CharStyle.Bold)!=0)
                    {
                        tokens.Add("bold");
                    }
                    if ((fci.Styles[i] & VA.Text.CharStyle.Italic)!=0)
                    {
                        tokens.Add("italic");
                    }
                    sb.Append(string.Join(",", tokens));
                    sb.Append(")");
                    fci.TargetFontDisplayNames.Add(sb.ToString());

                }
                else
                {
                    fci.TargetFontDisplayNames.Add(fci.TargetFontNames[i]);
                }
            }
            var first_page = fci.Document.Pages[1];
            this.compate_text_blocks(fci);
            this.compare_glyph_chart(fci);
            this.compare_glyphs(fci);
            first_page.Delete(1);
        }

        void compate_text_blocks(FontCompareInput fci)
        {
            double left = 1;
            double vs = 0.25;
            double cell1_w = 1.0;
            double cell1_h = 2.0;
            double cell1_top = 8.5;
            double cell2_h = 0.5;
            double cell2_w = 8.0;
            double cell_sep = 0.5;
            string labeltextcolor = "rgb(0,176,240)";

            var char1 = new VA.Text.CharacterCells();
            char1.Font = fci.LabelFontID;
            char1.Size = "30pt";
            char1.Color = labeltextcolor;

            var char2 = new VA.Text.CharacterCells();
            char2.Font = fci.LabelFontID;
            char2.Size = "16pt";
            char2.Color = labeltextcolor;

            var fmt1 = new VA.Shapes.FormatCells();
            fmt1.LineWeight = 0;
            fmt1.LinePattern = 0;
            fmt1.FillPattern = 0;

            var fmt2 = new VA.Shapes.FormatCells();
            fmt2.LineWeight = 0;
            fmt2.LinePattern = 0;
            fmt2.FillPattern = 0;

            var fmt3 = new VA.Shapes.FormatCells();
            fmt3.LineWeight = 0;
            fmt3.LinePattern = 0;
            fmt3.FillPattern = 0;

            var tb1 = new VA.Text.TextCells();
            tb1.VerticalAlign = 0;

            var para1 = new VA.Text.ParagraphCells();
            para1.HorizontalAlign = 2;

            var para2 = new VA.Text.ParagraphCells();
            para2.HorizontalAlign = 0;

            var char3 = new VA.Text.CharacterCells();

            var para3 = new VA.Text.ParagraphCells();
            para3.HorizontalAlign = 0;

            var tb3 = new VA.Text.TextCells();
            tb3.VerticalAlign = 0;

            foreach (string text in fci.TextBlocks)
            {
                var curpage = fci.Document.Pages.Add();
                foreach (var size in fci.FontSizeFormulas)
                {
                    var shape1 = curpage.DrawRectangle(left, cell1_top - cell1_h, left + cell1_w, cell1_top);
                    shape1.Text = string.Format("{0}", size);

                    var update1 = new VA.ShapeSheet.Update();
                    update1.SetFormulas(para1,0);
                    update1.SetFormulas(char1,0);
                    update1.SetFormulas(tb1);
                    update1.SetFormulas(fmt1);

                    update1.Execute(shape1);

                    double cell2_top = cell1_top;
                    for (int i = 0; i < fci.TargetFontNames.Count(); i++)
                    {
                        double cell2_bottom = cell2_top - cell2_h;
                        var fontname = fci.TargetFontDisplayNames[i];
                        double cell2_left = left + cell1_w + cell_sep;
                        var shape2 = curpage.DrawRectangle(cell2_left, cell2_bottom, cell2_left + cell2_w, cell2_top);
                        shape2.Text = string.Format("{0}", fontname);

                        double cell3_h = 2.0; // default height
                        var cell3_top = cell2_bottom;
                        var cell3_bottom = cell3_top - cell3_h;

                        var shape_3 = curpage.DrawRectangle(cell2_left, cell3_bottom, cell2_left + cell2_w, cell3_top);
                        shape_3.Text = text;


                        char3.Font = fci.TargetFontIDs[i];
                        char3.Size = size;
                        char3.Style = (int) (fci.Styles[i]);

                        var update3 = new VA.ShapeSheet.Update();

                        update3.SetFormulas(para3,0);
                        update3.SetFormulas(tb3);
                        update3.SetFormulas(char3,0);
                        update3.SetFormulas(fmt3);
                        update3.Execute(shape_3);


                        char2.Font = fci.TargetFontIDs[i];
                        var update2 = new VA.ShapeSheet.Update();

                        update2.SetFormulas(para2,0);

                        update2.SetFormulas(char2, 0);
                        update2.SetFormulas(fmt2);

                        update2.Execute(shape2);


                        shape_3.CellsU["Height"].FormulaU = "TEXTHEIGHT(TheText,TxtWidth)";
                        var cell3_real_size = new VA.Drawing.Size(shape_3.CellsU["Width"].get_Result(null),
                                                                  shape_3.CellsU["Height"].get_Result(null));
                        shape_3.CellsU["PinY"].FormulaU = (cell2_bottom - (cell3_real_size.Height / 2.0)).ToString();

                        cell2_top -= cell2_h + cell3_real_size.Height + vs;
                    }
                    cell1_top = cell2_top;
                }

                curpage.ResizeToFitContents( new VisioAutomation.Drawing.Size(1.0, 1.0));
            }
        }

        void compare_glyph_chart(FontCompareInput fci)
        {
            var text =
                "ABCDEFGHIJKLMNOPQRSTUVWXYZ\nabcdefghijklmnopqrstuvqxyz\n1234567890\n!@#$%^&*()\n`~_-+=[]{}<>\n\\|;:'\",./?";
            var text1_x_lines = text.Split('\n');

            var page = fci.Document.Pages.Add();
            var dom = new VA.DOM.Document();

            var dom_page = new VA.DOM.Page();

            dom.Pages.Add(dom_page);
            double cy = 8.0;
            for (int fi = 0; fi < fci.TargetFontNames.Count; fi++)
            {
                var font = fci.TargetFontNames[fi];

                var trect = new VA.Drawing.Rectangle(0, cy - 1.0, 10, cy);
                var xshape = new VA.DOM.Shape(this.rectmaster, this.stencilname, trect.Center);
                xshape.Cells.Width = trect.Size.Width;
                xshape.Cells.Height = trect.Size.Height;
                dom_page.Shapes.Add(xshape);

                xshape.Text = new VA.Text.Markup.TextElement(font);
                xshape.Cells.FillPattern = 0;
                xshape.Cells.ParaHorizontalAlign = 0;
                xshape.CharFontName = font;
                xshape.Cells.CharSize = "36pt";
                xshape.Cells.LinePattern = 0;
                xshape.Cells.LineWeight = 0;
                cy -= 2.0;

                for (int ri = 0; ri < text1_x_lines.Length; ri++)
                {
                    var curline = text1_x_lines[ri];
                    for (int c = 0; c < curline.Length; c++)
                    {
                        double x = 0 + (1.0) * c;
                        var rect = new VA.Drawing.Rectangle(x, cy, x + 0.5, cy + 0.5);

                        var shape = dom_page.Shapes.Drop(this.rectmaster, this.stencilname, rect);
                        shape.Text = new VA.Text.Markup.TextElement(curline[c].ToString());
                        shape.Cells.FillPattern = 0;
                        shape.Cells.CharSize = "18pt";
                        shape.Cells.LineColor = "rgb(230,230,230)";
                        shape.Cells.CharStyle = (int) (fci.Styles[fi]);
                    }
                    cy -= 1.0;
                }
            }
            dom.Render(page.Application);
            page.ResizeToFitContents( new VisioAutomation.Drawing.Size(1.0, 1.0));
        }

        private void compare_glyphs(FontCompareInput fci)
        {
            var colorints = new[]
                             {
                                    0x00B0F0, 0xff0000, 0x00B050
                             };


            var colors = colorints.Select(i => new VA.Drawing.ColorRGB(i)).Select(c => c.ToFormula()).ToArray();

            var text =
                "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvqxyz1234567890«»ßäüöĆćŚśŹźŻżĄąĘę!@#$%^&*()`~_-+=[]{}\\|;:'\",.<>/?";
            var texts = MainForm.Split(text, 10);

            var pages = fci.Document.Pages;
            double col_width = 3.0;
            double label_height = 0.5;
            double sample_height = col_width;
            var dom = new VA.DOM.Document();

            for (int i = 0; i < texts.Count; i++)
            {

                var dom_page = new VA.DOM.Page();
                dom.Pages.Add(dom_page);
                double cy = 8.0;
                for (int j = 0; j < texts[i].Count; j++)
                {
                    double title_bottom = cy - label_height;
                    for (int k = 0; k < fci.TargetFontNames.Count(); k++)
                    {
                        double left = 0 + (k * col_width);
                        double right = left + col_width;
                        double top = cy;
                        var rect = new VA.Drawing.Rectangle(left, title_bottom, right, top);
                        var char_sample_title_shape = dom_page.Shapes.Drop(this.rectmaster, this.stencilname, rect);
                        char_sample_title_shape.Text = new VA.Text.Markup.TextElement(fci.TargetFontDisplayNames[k]);
                        char_sample_title_shape.Cells.LinePattern = 0;
                        char_sample_title_shape.Cells.LineWeight = 0;
                        char_sample_title_shape.Cells.CharFont = fci.TargetFontIDs[k];
                        char_sample_title_shape.Cells.FillPattern = 0;

                    }
                    double overlay_title_left = fci.TargetFontNames.Count()*col_width;

                    double overlay_title_right = overlay_title_left + col_width;
                    var rectangle = new VA.Drawing.Rectangle(overlay_title_left, title_bottom, overlay_title_right, cy);
                    var overlay_title_shape = new VA.DOM.Shape(this.rectmaster, this.stencilname, rectangle);
                    dom_page.Shapes.Add(overlay_title_shape);
                    overlay_title_shape.Cells.FillPattern = 0;
                    overlay_title_shape.Cells.LinePattern = 0;
                    overlay_title_shape.Cells.LineWeight = 0;
                    overlay_title_shape.Text = new VA.Text.Markup.TextElement();
                    overlay_title_shape.Cells.CharTransparency = "0.3";
                    for (int k = 0; k < fci.TargetFontNames.Count(); k++)
                    {
                        var s = fci.TargetFontDisplayNames[k] + "\n";
                        var el = overlay_title_shape.Text.AddElement(s);
                        el.CharacterCells.Color = (new VA.Drawing.ColorRGB(colorints[k % colorints.Count()])).ToFormula();
                        el.CharacterCells.Font = fci.TargetFontIDs[k];
                    }
                    cy -= label_height;

                    var cur_char = texts[i][j].ToString();
                    var bigcharsize = "160pt";
                    for (int k = 0; k < fci.TargetFontNames.Count(); k++)
                    {
                        double w = col_width;
                        double x0 = 0 + (k * w);
                        double x1 = x0 + w;
                        double y0 = cy - sample_height;
                        double y1 = cy;
                        var rect = new VA.Drawing.Rectangle(x0, y0, x1, y1);
                        var char_sample_shape = dom_page.Shapes.Drop(this.rectmaster, this.stencilname, rect);
                        char_sample_shape.Text = new VA.Text.Markup.TextElement(cur_char);
                        char_sample_shape.Cells.LinePattern = 0;
                        char_sample_shape.Cells.LineWeight = 0;
                        char_sample_shape.Cells.CharSize = bigcharsize;
                        char_sample_shape.Cells.CharFont = fci.TargetFontIDs[k];
                        char_sample_shape.Cells.FillPattern = 0;
                        char_sample_shape.Cells.CharStyle = ((int) fci.Styles[k]).ToString();
                    }

                    for (int k = 0; k < fci.TargetFontNames.Count(); k++)
                    {
                        double overlay_left = fci.TargetFontNames.Count()*col_width;
                        double overlay_right = overlay_left + col_width;
                        var rect = new VA.Drawing.Rectangle(overlay_left, cy - col_width, overlay_right, cy);

                        var overlay_shape = dom_page.Shapes.Drop(this.rectmaster, this.stencilname, rect);
                        overlay_shape.Cells.LinePattern = 0;
                        overlay_shape.Cells.LineWeight = 0;
                        overlay_shape.Text = new VA.Text.Markup.TextElement(cur_char);
                        overlay_shape.Cells.LinePattern = 0;
                        overlay_shape.Cells.LineWeight = 0;
                        overlay_shape.Cells.CharSize = bigcharsize;
                        overlay_shape.Cells.CharFont = fci.TargetFontIDs[k];
                        overlay_shape.Cells.FillPattern = 0;
                        overlay_shape.Cells.CharTransparency = "0.7";
                        overlay_shape.Cells.CharColor = colors[k % colors.Count()];
                        overlay_shape.Cells.CharStyle = ((int)fci.Styles[k]).ToString();
                    }

                    cy -= sample_height;

                    cy -= 1.0; // extra spacing
                }
                //page.ResizeToFitContents(new VisioAutomation.Drawing.Size(1.0, 1.0));

            }
            dom.Render(fci.Document.Application);
        }

        public static List<List<T>> Split<T>(IEnumerable<T> source, int n)
        {
            return source
                .Select((x, i) => new { Index = i, Value = x })
                .GroupBy(x => x.Index / n)
                .Select(x => x.Select(v => v.Value).ToList())
                .ToList();
        }

        private void buttonRun_Click(object sender, EventArgs e)
        {
            this.FontCompare();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }

    }

    public class FontCompareInput
    {
        public List<string> TargetFontNames;
        public IVisio.Document Document;
        public string[] TextBlocks;
        public List<string> FontSizeFormulas;
        public List<int> TargetFontIDs;
        public int LabelFontID;
        public List<string> TargetFontDisplayNames;
        public List<VA.Text.CharStyle> Styles;
    }
}
