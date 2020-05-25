using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Extensions {
    public partial class ColourPicker : ComboBox {
        public class ColourInfo {
            public String Text { get; }
            public OlCategoryColor OutlookCategory { get; }
            public Color Colour { get; }

            public ColourInfo(OlCategoryColor category, Color colour, String name = "") {
                this.Text = string.IsNullOrEmpty(name) ? OutlookOgcs.Categories.FriendlyCategoryName(category) : name;
                this.Colour = colour;
                this.OutlookCategory = category;
            }
        }

        public enum ColourType {
            OutlookStandardColours,
            OutlookCategoryColours
        }

        public ColourPicker() {
            DropDownStyle = ComboBoxStyle.DropDownList;
            DrawMode = DrawMode.OwnerDrawFixed;
            DrawItem += ColourPicker_DrawItem;
        }

        public void AddColourItems(ColourType? type) {
            if (type == null)
                Items.Clear();
            if (type == null || type == ColourType.OutlookCategoryColours)
                AddCategoryColours();
            if (type == null || type == ColourType.OutlookStandardColours)
                addStandardColours();
        }

        private void addStandardColours() {
            foreach (KeyValuePair<OlCategoryColor, Color> colour in OutlookOgcs.CategoryMap.Colours) {
                Items.Add(new ColourInfo(colour.Key, colour.Value));
            }
        }

        public void AddCategoryColours() {
            Items.AddRange(OutlookOgcs.Calendar.Categories.DropdownItems().ToArray());
        }

        public void ColourPicker_DrawItem(object sender, System.Windows.Forms.DrawItemEventArgs e) {
            ComboBox cbColour = sender as ComboBox;
            if (e == null || e.Index < 0 || e.Index >= cbColour.Items.Count)
                return;

            // Get the colour
            ColourInfo colour = (ColourInfo)Items[e.Index];
            ComboboxColor.DrawComboboxItemColour(cbColour, new SolidBrush(colour.Colour), colour.Text, e);
        }

        public new ColourInfo SelectedItem {
            get { return (ColourInfo)base.SelectedItem; }
            set { base.SelectedItem = value; }
        }
    }

    public class DataGridViewColourComboBoxColumn : DataGridViewColumn {
        public DataGridViewColourComboBoxColumn() : base(new DataGridViewColourComboBoxCell()) {
        }

        public override DataGridViewCell CellTemplate {
            get {
                return base.CellTemplate;
            }
            set {
                // Ensure that the cell used for the template is a DataGridViewColourComboBoxCell.
                if (value != null && !value.GetType().IsAssignableFrom(typeof(DataGridViewColourComboBoxCell))) {
                    throw new InvalidCastException("Must be a DataGridViewColourComboBoxCell");
                }
                base.CellTemplate = value;
            }
        }
    }

    public class DataGridViewColourComboBoxCell : DataGridViewTextBoxCell {

        public DataGridViewColourComboBoxCell() : base() { }

        public override void InitializeEditingControl(int rowIndex, object initialFormattedValue, DataGridViewCellStyle dataGridViewCellStyle) {
            base.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle);
            ComboboxColor ctl = DataGridView.EditingControl as ComboboxColor;
            if (this.RowIndex >= 0) {
                if (this.Value == null)
                    ctl.SelectedItem = (ColourPicker.ColourInfo)this.DefaultNewRowValue;
                else {
                    String currentText = this.Value.ToString();
                    if (ctl.Items.Count == 0) {
                        ctl.PopulateDropdownItems();
                    }
                    this.Value = currentText;
                    foreach (ColourPicker.ColourInfo ci in Forms.ColourMap.OutlookComboBox.Items) {
                        if (ci.Text == (String)this.Value) {
                            ctl.SelectedValue = ci;
                            break;
                        }
                    }
                }
            }
        }

        public override Type EditType {
            get {
                return typeof(ComboboxColor);
            }
        }

        public override Type ValueType {
            get {
                return typeof(ColourPicker.ColourInfo);
            }
        }

        public override object DefaultNewRowValue {
            get {
                if (Forms.ColourMap.OutlookComboBox.Items.Count > 0)
                    return (Forms.ColourMap.OutlookComboBox.Items[1] as ColourPicker.ColourInfo).Text;
                else
                    return String.Empty;
            }
        }

        protected override void Paint(System.Drawing.Graphics graphics, System.Drawing.Rectangle clipBounds, System.Drawing.Rectangle cellBounds, int rowIndex, System.Windows.Forms.DataGridViewElementStates elementState, object value, object formattedValue, string errorText, System.Windows.Forms.DataGridViewCellStyle cellStyle, System.Windows.Forms.DataGridViewAdvancedBorderStyle advancedBorderStyle, System.Windows.Forms.DataGridViewPaintParts paintParts) {
            //Paint inactive cells
            //base.Paint(graphics, clipBounds, cellBounds, rowIndex, elementState, value, formattedValue, errorText, cellStyle, advancedBorderStyle, paintParts);

            int indexItem = rowIndex;
            if (indexItem < 0)
                return;

            foreach (ColourPicker.ColourInfo ci in Forms.ColourMap.OutlookComboBox.Items) {
                if (ci.Text == this.Value.ToString()) {
                    Brush boxBrush = new SolidBrush(ci.Colour);
                    Brush textBrush = SystemBrushes.WindowText;
                    Extensions.ComboboxColor.DrawComboboxItemColour(true, boxBrush, textBrush, this.Value.ToString(), graphics, cellBounds);
                    break;
                }
            }
        }
    }

    public class ComboboxColor : ComboBox, IDataGridViewEditingControl {
        DataGridView dataGridView;
        private bool valueChanged = false;
        int rowIndex;

        public ComboboxColor() {
            PopulateDropdownItems();

            this.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DrawItem += new DrawItemEventHandler(ComboboxColor_DrawItem);
            this.SelectedIndexChanged += new EventHandler(ComboboxColor_SelectedIndexChanged);
        }

        public void PopulateDropdownItems() {
            Dictionary<Extensions.ColourPicker.ColourInfo, String> cbItems = new Dictionary<Extensions.ColourPicker.ColourInfo, String>();
            foreach (Extensions.ColourPicker.ColourInfo ci in Forms.ColourMap.OutlookComboBox.Items) {
                cbItems.Add(ci, ci.Text);
            }
            this.DataSource = new BindingSource(cbItems, null);
            this.DisplayMember = "Value";
            this.ValueMember = "Key";
        }

        public object EditingControlFormattedValue {
            get {
                return this.FormatString;
            }
            set {
                if (value is String) {
                    try {
                        this.FormatString = (string)value;
                    } catch {
                        this.FormatString = string.Empty;
                    }
                }
            }
        }

        public object GetEditingControlFormattedValue(DataGridViewDataErrorContexts context) {
            return EditingControlFormattedValue;
        }

        public void ApplyCellStyleToEditingControl(DataGridViewCellStyle dataGridViewCellStyle) {
            this.Font = dataGridViewCellStyle.Font;
            this.ForeColor = dataGridViewCellStyle.ForeColor;
            this.BackColor = dataGridViewCellStyle.BackColor;
        }

        public int EditingControlRowIndex {
            get {
                return rowIndex;
            }
            set {
                rowIndex = value;
            }
        }

        public bool EditingControlWantsInputKey(Keys key, bool dataGridViewWantsInputKey) {
            switch (key & Keys.KeyCode) {
                case Keys.Left:
                case Keys.Up:
                case Keys.Down:
                case Keys.Right:
                case Keys.Home:
                case Keys.End:
                case Keys.PageDown:
                case Keys.PageUp:
                    return true;
                default:
                    return !dataGridViewWantsInputKey;
            }
        }

        public void PrepareEditingControlForEdit(bool selectAll) {
        }

        public bool RepositionEditingControlOnValueChange {
            get {
                return false;
            }
        }

        public DataGridView EditingControlDataGridView {
            get {
                return dataGridView;
            }
            set {
                dataGridView = value;
            }
        }

        public bool EditingControlValueChanged {
            get {
                return valueChanged;
            }
            set {
                valueChanged = value;
            }
        }

        public Cursor EditingPanelCursor {
            get {
                return base.Cursor;
            }
        }

        void ComboboxColor_DrawItem(object sender, DrawItemEventArgs e) {
            ComboBox cbColour = sender as ComboBox;
            int indexItem = e.Index;
            if (indexItem < 0 || indexItem >= cbColour.Items.Count)
                return;

            KeyValuePair<Extensions.ColourPicker.ColourInfo, String> kvp = (KeyValuePair<Extensions.ColourPicker.ColourInfo, String>)cbColour.Items[indexItem];
            if (kvp.Key != null) {
                // Get the colour
                OlCategoryColor olColour = kvp.Key.OutlookCategory;
                Brush brush = new SolidBrush(OutlookOgcs.CategoryMap.RgbColour(olColour));

                DrawComboboxItemColour(cbColour, brush, kvp.Value, e);
            }
        }

        void ComboboxColor_SelectedIndexChanged(object sender, EventArgs e) {
            if (dataGridView.SelectedCells != null && dataGridView.SelectedCells.Count > 0)
                dataGridView.SelectedCells[0].Value = this.Text;
        }

        public static void DrawComboboxItemColour(ComboBox cbColour, Brush boxColour, String itemDescription, DrawItemEventArgs e) {
            try {
                e.Graphics.FillRectangle(new SolidBrush(cbColour.BackColor), e.Bounds);
                e.DrawBackground();
                Boolean comboEnabled = cbColour.Enabled;

                // Write colour name
                Boolean highlighted = (e.State & DrawItemState.Selected) != DrawItemState.None;
                Brush brush = comboEnabled ? SystemBrushes.WindowText : SystemBrushes.InactiveCaptionText;
                if (highlighted)
                    brush = comboEnabled ? SystemBrushes.HighlightText : SystemBrushes.InactiveCaptionText;

                DrawComboboxItemColour(comboEnabled, boxColour, brush, itemDescription, e.Graphics, e.Bounds);

                // Draw the focus rectangle if appropriate
                if ((e.State & DrawItemState.NoFocusRect) == DrawItemState.None)
                    e.DrawFocusRectangle();
            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }
        }

        public static void DrawComboboxItemColour(Boolean comboEnabled, Brush boxColour, Brush textColour, String itemDescription, Graphics graphics, Rectangle cellBounds) {
            try {
                // Draw colour box
                Rectangle colourbox = new Rectangle();
                colourbox.X = cellBounds.X + 2;
                colourbox.Y = cellBounds.Y + 2;
                colourbox.Height = cellBounds.Height - 5;
                colourbox.Width = 18;
                graphics.FillRectangle(boxColour, colourbox);
                graphics.DrawRectangle(comboEnabled ? SystemPens.WindowText : SystemPens.InactiveBorder, colourbox);

                int textX = cellBounds.X + colourbox.X + colourbox.Width + 2;

                graphics.DrawString(itemDescription, Control.DefaultFont, textColour,
                    /*cellBounds.X*/ +colourbox.X + colourbox.Width + 2,
                    cellBounds.Y + ((cellBounds.Height - Control.DefaultFont.Height) / 2));

            } catch (System.Exception ex) {
                OGCSexception.Analyse(ex);
            }
        }
    }
}
